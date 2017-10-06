using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using DocumentFormat.OpenXml.Extensions;
using System.IO;
using DocumentFormat.OpenXml;
using System.Collections;
using System.Reflection;
using FormulaExcel;
using System.Text.RegularExpressions;
using ExcelFormulaParser;
using System.Globalization;

namespace DocumentFormat.OpenXml.ReportBuilder
{
    public static class ReportBuilderXLS
    {
        public static UInt32 ProgressValue = 0, MaximumCount = 0;
        private static UInt32 serviceRow = 1;
        private static SharedStringTable sharedStringTable { get; set; }

        public static MemoryStream GenerateReport(Dictionary<string, IList> dictionary, string templateFileName, bool protectionSheet)
        {
            return GenerateReport(dictionary, templateFileName, null, protectionSheet);
        }

        public static MemoryStream GenerateReport(DataSet dataSet, string Template, bool protectionSheet)
        {
            MemoryStream ms = new MemoryStream();
            ms = SpreadsheetReader.Copy(Template);

            return GenerateReport(dataSet, ms, null, protectionSheet);
        }

        public static MemoryStream GenerateReport(Dictionary<string, IList> dictionary, string templateFileName)
        {
            return GenerateReport(dictionary, templateFileName, null, false);
        }

        public static MemoryStream GenerateReport(DataSet dataSet, string Template)
        {
            MemoryStream ms = new MemoryStream();
            ms = SpreadsheetReader.Copy(Template);

            return GenerateReport(dataSet, ms, null, false);
        }

        public static MemoryStream GenerateReport(DataSet dataSet, MemoryStream Template)
        {
            return GenerateReport(dataSet, Template, null, false);
        }


        public static MemoryStream GenerateReport(Dictionary<string, IList> dictionary, string templateFileName, string OutFile, bool protectionSheet)
        {
            MemoryStream ms = new MemoryStream();
            ms = SpreadsheetReader.Copy(templateFileName);

            DataSet dataSet = new DataSet();
            foreach (var obj in dictionary.Where(w => w.Value.Count > 0))
            {
                /*    DataTable dataTable = new DataTable(obj.Key);
                    if (obj.Value.Count > 0)
                    {
                  //     var dt =  obj.Value.Cast<Object>().Ext_ToDataTable(obj.Key);

                        object row = dictionary[obj.Key][0];
                        Type type = row.GetType();
                        PropertyInfo[] oProps = type.GetProperties();
                        foreach (PropertyInfo pi in oProps)
                        {
                            Type colType = pi.PropertyType;

                            if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                            {
                                colType = colType.GetGenericArguments()[0];
                            }

                            dataTable.Columns.Add(new DataColumn(pi.Name, colType));
                        }

                        foreach (Object Data in obj.Value)
                        {
                            Type personType = Data.GetType();
                            PropertyInfo[] pis = personType.GetProperties();
                            DataRow newRow = dataTable.NewRow();

                            foreach (PropertyInfo pin in pis)
                            {
                                newRow[pin.Name] = pin.GetValue(Data, null) ?? DBNull.Value; 
                            }
                            dataTable.Rows.Add(newRow);

                       }
                    }
                    dataTable.AcceptChanges();*/

                dataSet.Tables.Add(obj.Value.Cast<Object>().Ext_ToDataTable(obj.Key));

            }

            return GenerateReport(dataSet, ms, OutFile, protectionSheet);
        }

        public static DataTable Ext_ToDataTable<T>(this IEnumerable<T> varlist, string name)
        {
            DataTable dtReturn = new DataTable(name);

            // column names 
            PropertyInfo[] oProps = null;
            FieldInfo[] oField = null;
            if (varlist == null) return dtReturn;

            foreach (T rec in varlist)
            {
                // Use reflection to get property names, to create table, Only first time, others will follow 
                if (oProps == null)
                {
                    oProps = rec.GetType().GetProperties();
                    foreach (PropertyInfo pi in oProps)
                    {
                        Type colType = pi.PropertyType;

                        if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                        {
                            colType = colType.GetGenericArguments()[0];
                        }

                        dtReturn.Columns.Add(new DataColumn(pi.Name, colType));
                    }
                    oField = rec.GetType().GetFields();
                    foreach (FieldInfo fieldInfo in oField)
                    {
                        Type colType = fieldInfo.FieldType;

                        if ((colType.IsGenericType) &&
                             (colType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                        {
                            colType = colType.GetGenericArguments()[0];
                        }

                        dtReturn.Columns.Add(new DataColumn(fieldInfo.Name, colType));
                    }
                }

                DataRow dr = dtReturn.NewRow();

                if (oProps != null)
                {
                    foreach (PropertyInfo pi in oProps)
                    {
                        dr[pi.Name] = pi.GetValue(rec, null) ?? DBNull.Value;
                    }
                }
                if (oField != null)
                {
                    foreach (FieldInfo fieldInfo in oField)
                    {
                        dr[fieldInfo.Name] = fieldInfo.GetValue(rec) ?? DBNull.Value;
                    }
                }
                dtReturn.Rows.Add(dr);
            }
            return dtReturn;
        }


       // if( cell.CellFormula.FormulaType.Value == CellFormulaValues.Shared ) cell.CellFormula.

        public class SharedList
        {
            public UInt32 idx;
            public string Ref;
            public string formula;
            public List<Cell> lCell = new List<Cell>();
        }

        private static void UpdateSharedFormula(SpreadsheetDocument RecultDoc)
        {
            List<SharedList> sharedList = new List<SharedList>();
            foreach (Sheet curSheet in RecultDoc.WorkbookPart.Workbook.Sheets)
            {
                WorksheetPart RecultWorkSheetPart = SpreadsheetReader.GetWorksheetPartByName(RecultDoc, curSheet.Name.Value);
                SheetData sheetData = RecultWorkSheetPart.Worksheet.Elements<SheetData>().First();
                List<Row> selctRow = new List<Row>(sheetData.Elements<Row>());
                foreach (Row curRow in selctRow)
                {
                    foreach (Cell curCell in curRow.Elements<Cell>().Where(c => c.CellFormula != null)
                                                                      .Where(c => c.CellFormula.FormulaType != null)
                                                                        .Where(c => c.CellFormula.FormulaType.Value == CellFormulaValues.Shared))
                    {
                        if (curCell.CellFormula.Reference != null)
                        {
                            SharedList sl = new SharedList();
                            sl.idx = curCell.CellFormula.SharedIndex;
                            sl.Ref = curCell.CellFormula.Reference.Value;
                            sl.formula = curCell.CellFormula.Text;
                            foreach (Row chRow in selctRow)
                            {
                                foreach (Cell chCell in chRow.Elements<Cell>().Where(c => c.CellFormula != null)
                                                                                .Where(c => c.CellFormula.FormulaType != null)
                                                                                 .Where(c => c.CellFormula.SharedIndex == sl.idx))
                                {
                                    sl.lCell.Add(chCell);
                                }
                            }
                            sharedList.Add(sl);
                        }
                    }
                }       
            }

            foreach (SharedList sl in sharedList)
            {
                string fRef = FirstRefFromRange(sl.Ref);
                string lRef = LastRefFromRange(sl.Ref);
                string fColinRef = SpreadsheetReader.ColumnFromReference(fRef);
                string lColinRef = SpreadsheetReader.ColumnFromReference(lRef);
                if(fColinRef == lColinRef)
                {
                    int offset = 0;
                    ExcelFormula excelFormula = new ExcelFormula("=" + sl.formula);
                    
                    foreach (Cell cell in sl.lCell)
                    {
                        string fStr = "";
                        for (int i = 0; i < excelFormula.Count; i++)
                        {

                            ExcelFormulaToken token = excelFormula[i];
                            string tokenValue = token.Value;

                            if (token.Subtype == ExcelFormulaTokenSubtype.Start) tokenValue += "(";
                            if (token.Subtype == ExcelFormulaTokenSubtype.Text) tokenValue = '\"' + tokenValue.Replace("\"", "\"\"") + '\"';
                            if (token.Subtype == ExcelFormulaTokenSubtype.Reference)
                            {
                                UInt32 row = SpreadsheetReader.RowFromReference(tokenValue);
                                string col = SpreadsheetReader.ColumnFromReference(tokenValue);
                                tokenValue = col + (row + offset);
                            }
                            if (token.Subtype == ExcelFormulaTokenSubtype.Range)
                            {
                                string fr = FirstRefFromRange(tokenValue);
                                string lr = LastRefFromRange(tokenValue);
                                tokenValue = SpreadsheetReader.ColumnFromReference(fr) + (SpreadsheetReader.RowFromReference(fr) + offset) + ":" + SpreadsheetReader.ColumnFromReference(lr) + (SpreadsheetReader.RowFromReference(lr) + offset);
                            }
                            if (token.Subtype == ExcelFormulaTokenSubtype.Stop) tokenValue = ")";
                            fStr += tokenValue;
                        }
                        cell.CellFormula.Remove();
                        cell.CellFormula = new CellFormula(fStr);

                      //  cell.CellFormula.FormulaType.InnerText = "e";
                     //   cell.CellFormula.Text = fStr;
                        offset++;

                    }
                }
            }


        }

        //Создание отчета
        public static MemoryStream GenerateReport(DataSet dataSet, MemoryStream ms, string OutFile, bool protectionSheet)
        {
            SpreadsheetDocument RecultDoc = SpreadsheetDocument.Open(ms, true);
            sharedStringTable = RecultDoc.WorkbookPart.SharedStringTablePart.SharedStringTable;

            UpdateSharedFormula( RecultDoc);

            foreach (Sheet curSheet in RecultDoc.WorkbookPart.Workbook.Sheets)
            {
                WorksheetPart RecultWorkSheetPart = SpreadsheetReader.GetWorksheetPartByName(RecultDoc, curSheet.Name.Value);
                SheetData sheetData = RecultWorkSheetPart.Worksheet.Elements<SheetData>().First();

                // Получаем имена областей DefinedNames принадлежащих к текущему Sheet
                List<DefinedName> dnList = new List<DefinedName>();
                if (RecultDoc.WorkbookPart.Workbook.Elements<DefinedNames>().Count() > 0)
                {
                    dnList = RecultDoc.WorkbookPart.Workbook.Elements<DefinedNames>().First().Elements<DefinedName>()
                                                                 .Where(d => d.Text.IndexOf("#REF!") < 0 && curSheet.Name.Value == d.Text.Substring(0, d.Text.IndexOf("!")).TrimEnd('\'').TrimStart('\'')).OrderBy(d => FirstRowInRange(d.Text.Substring(d.Text.IndexOf("!") + 1))).ToList();
                }

                //Выбираем все строки с SheetData которые не принадлежат ни одному из DefinedName
                List<Row> selctRow = new List<Row>(sheetData.Elements<Row>());
                foreach (DefinedName DN in dnList)
                {
                    string Range = GetRangeInDefName(DN);
                    selctRow = selctRow.Where(r => r.RowIndex < FirstRowInRange(Range) || r.RowIndex > LastRowInRange(Range)).ToList();
                }

                if (sharedStringTable != null)
                {
                    int i = sharedStringTable.Count() - 1;
                    foreach (Row curRow in selctRow)
                    {
                        foreach (Cell curCell in curRow.Elements<Cell>().Where(c => c.DataType != null && c.DataType.Value == CellValues.SharedString))
                        {
                            var SharedString = sharedStringTable.ChildElements[Int32.Parse(curCell.CellValue.InnerText)].CloneNode(true);
                            var parcing_Xml = ParcingCellValue(dataSet, null, SharedString.InnerXml);
                            if (String.Compare(parcing_Xml, SharedString.InnerXml) > 0)
                            {
                                SharedString.InnerXml = parcing_Xml;
                                sharedStringTable.Append(SharedString);

                                curCell.CellValue = new CellValue((++i).ToString());
                                curCell.DataType.Value = CellValues.SharedString;
                            }
                        }
                    }
                    sharedStringTable.Save();
                }

                //Заполняем выбраные ячейки selctRow даными с dataSet согласно формулам
                foreach (Row curRow in selctRow)
                {
                    foreach (Cell curCell in curRow.Elements<Cell>().Where(c => c.CellFormula != null))
                    {
                        SetValueToCellFormula(dataSet, null, curCell);
                    }
                }

                foreach (DefinedName definedName in dnList)
                {
                    DataTable dataTable = GetTableByName(dataSet, definedName.Name);
                    if (dataTable != null)
                    {
                        string oldRange = GetRangeInDefName(definedName);
                        UInt32 GlobalOffSet = SetValueFromDataTable(dataTable, dnList, RecultDoc);
                        selctRow = sheetData.Elements<Row>().Where(r => r.RowIndex < FirstRowInRange(GetRangeInDefName(definedName)) || r.RowIndex > LastRowInRange(GetRangeInDefName(definedName))).ToList();
                        UpdateFormulaWithoutRange(oldRange, selctRow, GlobalOffSet);
                    }
                }

                if (protectionSheet) //Защита листа случайно сгенерированым паролем
                {
                    SheetProtection prch = new SheetProtection();
                    prch.Password = Convert.ToString(new Random().Next(245, 12457));
                    prch.Sheet = true;
                    RecultWorkSheetPart.Worksheet.InsertAfter(prch, RecultWorkSheetPart.Worksheet.Elements<SheetData>().First());
                }

            }
            ClearSpreadsheetDocument(RecultDoc);

            RecultDoc.Close();

            if (OutFile != null)
                if (OutFile.Length > 0) SpreadsheetWriter.StreamToFile(OutFile, ms);

            return ms;
        }

        private static string ParcingCellValue(DataSet dataSet, DataRow dataRow, string cell_value)
        {
            string RangeName = "";

            if (dataRow != null)
            {
                RangeName = dataRow.Table.TableName;
            }

            foreach (DataTable Table in dataSet.Tables)
            {
                foreach (DataColumn Dcol in Table.Columns)
                {
                    string val;
                    string teg = "{{" + Table.TableName + "_" + Dcol.ColumnName + "}}";

                    var pattern = @"(\W|^)(" + teg + @")(\W|$)";

                    if (Regex.IsMatch(cell_value, pattern))
                    {
                        bool str = (Dcol.DataType.Name == "String" || Dcol.DataType.Name == "DateTime");
                        bool Dec = (Dcol.DataType.Name == "Decimal" || Dcol.DataType.Name == "Double");
                        bool Int = (Dcol.DataType.Name == "Int32" || Dcol.DataType.Name == "Int64");

                        bool nullDB;
                        if (RangeName == Table.TableName)
                        {
                            val = dataRow[Dcol.ColumnName].ToString();
                            nullDB = Convert.IsDBNull(dataRow[Dcol.ColumnName]);
                        }
                        else
                        {
                            val = Table.Rows[0][Dcol.ColumnName].ToString();
                            nullDB = Convert.IsDBNull(Table.Rows[0][Dcol.ColumnName]);
                        }

                        if (Dec)
                        {
                            val = val.Replace(',', '.');
                        }

                        if ((Dec || Int) && val.Length == 0)
                        {
                            val = "0";
                        }

                        cell_value = Regex.Replace(cell_value, pattern, "${1}" + val + "${3}");

                    }
                }
            }

            return cell_value;
        }


        private static void UpdateFormulaWithoutRange(String OldRange, List<Row> selctRow, UInt32 offset)
        {
            // Обновляем формули по всем ячейкам листа которые не принадлежат ни одному из DefinedName

            UInt32 frowDN = FirstRowInRange(OldRange);
            UInt32 endrowDN = LastRowInRange(OldRange);

            foreach (Row curRow in selctRow)
            {
                foreach (Cell cell in curRow.Elements<Cell>().Where(c => c.CellFormula != null))
                {
                    ExcelFormula excelFormula = new ExcelFormula("=" + cell.CellFormula.Text);
                    string fStr = "";

                    for (int i = 0; i < excelFormula.Count; i++)
                    { 

                        ExcelFormulaToken token = excelFormula[i];
                        string tokenValue = token.Value;

                        if (token.Subtype == ExcelFormulaTokenSubtype.Start) tokenValue += "(";
                        if (token.Subtype == ExcelFormulaTokenSubtype.Text) tokenValue = '\"' + tokenValue.Replace("\"", "\"\"") + '\"';
                        if (token.Subtype == ExcelFormulaTokenSubtype.Reference)
                        {
                            UInt32 row = SpreadsheetReader.RowFromReference(tokenValue);
                            if (row > frowDN) tokenValue = SpreadsheetReader.ColumnFromReference(tokenValue) + (row + offset);
                        }
                        if (token.Subtype == ExcelFormulaTokenSubtype.Range)
                        {
                            string fr = FirstRefFromRange(tokenValue);
                            string lr = LastRefFromRange(tokenValue);
                            if (SpreadsheetReader.RowFromReference(fr) > frowDN)
                                tokenValue = SpreadsheetReader.ColumnFromReference(fr) + (SpreadsheetReader.RowFromReference(fr) + offset) + ":" + SpreadsheetReader.ColumnFromReference(lr) + (SpreadsheetReader.RowFromReference(lr) + offset);
                        }
                        if (token.Subtype == ExcelFormulaTokenSubtype.Stop) tokenValue = ")";
                        fStr += tokenValue;
                    }
                    cell.CellFormula.Text = fStr;
                }
            }

        }

        private static UInt32 SetValueFromDataTable(DataTable dataTable, List<DefinedName> listDefNameInSheet, SpreadsheetDocument RecultDoc)
        {
            UInt32 newOffset = 0;
            string RangeName = dataTable.TableName;
            DefinedName curDefinedName = SpreadsheetReader.GetDefinedName(RecultDoc, RangeName);
         
            string definedName = curDefinedName.Text;
            string WorkSheetName = definedName.Substring(0, definedName.IndexOf("!")).TrimEnd('\'').TrimStart('\'');
            string Range = definedName.Substring(definedName.IndexOf("!") + 1);
            UInt32 startRowInRange = FirstRowInRange(Range);
            UInt32 endRowInRange = LastRowInRange(Range);
            UInt32 CountRangeData = (endRowInRange - startRowInRange + 1) - serviceRow;
            int countRows = dataTable.Rows.Count;
            
            WorksheetPart RecultWorkSheetPart = SpreadsheetReader.GetWorksheetPartByName(RecultDoc, WorkSheetName);
            SheetData sheetData = RecultWorkSheetPart.Worksheet.GetFirstChild<SheetData>();

            List<Row> ListRowInRange = RecultWorkSheetPart.Worksheet.Elements<SheetData>().First().Elements<Row>()
                                                              .Where(r => r.RowIndex >= startRowInRange && r.RowIndex < endRowInRange).ToList();
            List<Row> ListRowBottom = RecultWorkSheetPart.Worksheet.Elements<SheetData>().First().Elements<Row>()
                                                              .Where(r => r.RowIndex > endRowInRange).ToList();
            List<Row> ServiceRow = RecultWorkSheetPart.Worksheet.Elements<SheetData>().First().Elements<Row>()
                                                              .Where(r => r.RowIndex == endRowInRange).ToList();

           //Создаем список ячеек которые нуждаются в обновлении формул
            List<Cell> updateFormulaInRange = new List<Cell>();
            foreach (Row row in ListRowInRange)
            {
                foreach (Cell cell in row.Elements<Cell>().Where(c => c.CellFormula != null))
                {
                    if (new ExcelFormula("=" + cell.CellFormula.Text).Where<ExcelFormulaToken>(t => t.Subtype == ExcelFormulaTokenSubtype.Range || t.Subtype == ExcelFormulaTokenSubtype.Reference).Count() > 0)
                        updateFormulaInRange.Add(cell);
                }
            }

            List<Cell> updateFormulaInServiceRow = new List<Cell>();
            foreach (Row row in ServiceRow)
            {
                foreach (Cell cell in row.Elements<Cell>().Where(c => c.CellFormula != null))
                {
                    if (new ExcelFormula("=" + cell.CellFormula.Text).Where<ExcelFormulaToken>(t => t.Subtype == ExcelFormulaTokenSubtype.Range || t.Subtype == ExcelFormulaTokenSubtype.Reference).Count() > 0)
                        updateFormulaInServiceRow.Add(cell);
                }
            }

            List<MergeCell> ListMCellInRange = new List<MergeCell>();
            List<MergeCell> ListMCellBottom = new List<MergeCell>();
            List<MergeCell> ListMCellInServiceRow = new List<MergeCell>();
            MergeCells mergeCells = RecultWorkSheetPart.Worksheet.GetFirstChild<MergeCells>();
            if (mergeCells != null)
            {
                ListMCellInRange = RecultWorkSheetPart.Worksheet.Elements<MergeCells>().First().Elements<MergeCell>()
                                                            .Where(r => SpreadsheetReader.RowFromReference(FirstRefFromRange(r.Reference.Value)) >= startRowInRange && SpreadsheetReader.RowFromReference(LastRefFromRange(r.Reference.Value)) <= endRowInRange - serviceRow).ToList();
                ListMCellBottom = RecultWorkSheetPart.Worksheet.Elements<MergeCells>().First().Elements<MergeCell>()
                                                            .Where(r => SpreadsheetReader.RowFromReference(FirstRefFromRange(r.Reference.Value)) > endRowInRange).ToList();
                ListMCellInServiceRow = RecultWorkSheetPart.Worksheet.Elements<MergeCells>().First().Elements<MergeCell>()
                                                                    .Where(r => SpreadsheetReader.RowFromReference(FirstRefFromRange(r.Reference.Value)) == endRowInRange).ToList();
            }
            List<DefinedName> ListDefNameBottom = listDefNameInSheet.Where(dn => SpreadsheetReader.RowFromReference(FirstRefFromRange(dn.Text.Substring(dn.Text.IndexOf("!") + 1))) > endRowInRange).ToList();

            if (ListRowBottom.Count() > 0) DeleteRows(RecultWorkSheetPart, startRowInRange, ListRowBottom.Last().RowIndex.Value - startRowInRange + 1);
            else DeleteRows(RecultWorkSheetPart, startRowInRange, CountRangeData + serviceRow);

            List<DataRow> ListDataRow = dataTable.Rows.Cast<DataRow>().Where(r => r.RowState != DataRowState.Deleted).ToList();
            MaximumCount = (uint)ListDataRow.Count;

            UInt32 CountAddRows = (CountRangeData * MaximumCount) /*+ (UInt32)ServiceRow.Count()*/ - CountRangeData;

            UInt32 InsideOffset = startRowInRange;
            foreach (DataRow dataRow in ListDataRow)
            {  
                UInt32 Offset = (InsideOffset - startRowInRange);
                AppendRows(dataRow, ListRowInRange, ListMCellInRange, updateFormulaInRange, RecultWorkSheetPart, Offset, CountAddRows);
                InsideOffset += CountRangeData;
                ++ProgressValue;
            }

            newOffset = (InsideOffset - startRowInRange - CountRangeData);

            //Обновляем ссилки начала и конца рабочих областей DefinedNames
            curDefinedName.Text = curDefinedName.Text.Substring(0, curDefinedName.Text.LastIndexOf('$') + 1) + InsideOffset.ToString();//Текущая облать DefinedName
 
            //Обновляем ниже стоящие DefinedNames текущего листа Sheet
            foreach (DefinedName CurdefName in ListDefNameBottom)
            {
                string strDefText = CurdefName.Text;
                string range = strDefText.Substring(strDefText.IndexOf("!") + 1);

                string firstRefInRange = FirstRefFromRange(range);
                string lastRefInRange = LastRefFromRange(range);
                UInt32 firstRow = SpreadsheetReader.RowFromReference(firstRefInRange);
                UInt32 lastRow = SpreadsheetReader.RowFromReference(lastRefInRange);
                UInt32 offset = (MaximumCount * CountRangeData) - CountRangeData;

                CurdefName.Text = strDefText.Substring(0, strDefText.IndexOf("!") + 1) + '$' + SpreadsheetReader.ColumnFromReference(firstRefInRange) + '$' + (firstRow + offset).ToString() + ":$" + SpreadsheetReader.ColumnFromReference(lastRefInRange) + '$' + (lastRow + offset).ToString();
            }

            //Добавляем в конец листа Сервисную строку текущего Ренджа
            AppendServiceRow(ServiceRow, ListMCellInServiceRow, updateFormulaInServiceRow, curDefinedName, RecultWorkSheetPart, RecultDoc, newOffset, CountAddRows, countRows);    
            //Добавляем в конец листа все что ниже конца текущего Ренджа
            AppendRows(null, ListRowBottom, ListMCellBottom, updateFormulaInRange, RecultWorkSheetPart, newOffset, CountAddRows);

            return newOffset;
        }

        private static void AppendServiceRow(List<Row> ServiceRow, List<MergeCell> ListMCellInServiceRow, List<Cell> updateFormulaInServiceRow, DefinedName curDefinedName, WorksheetPart RecultWorkSheetPart, SpreadsheetDocument RecultDoc, uint Offset, UInt32 CountAddRows, int countRows)
        {
            UInt32 CountRangeData = Convert.ToUInt32(ServiceRow.Count());
            SheetData sheetData = RecultWorkSheetPart.Worksheet.Elements<SheetData>().First();
          //  string definedName = curDefinedName.Text;
            string range = GetRangeInDefName(curDefinedName);
            UInt32 firstRow = SpreadsheetReader.RowFromReference(FirstRefFromRange(range));
            UInt32 lastRow = SpreadsheetReader.RowFromReference(LastRefFromRange(range)) - serviceRow;

            foreach (Row RangeRow in ServiceRow) //Добавляем строки выбраные по указаному DefinedName 
            {
                Row r = (Row)RangeRow.Clone();
                r.RowIndex = r.RowIndex + Offset;
                foreach (Cell cell in r.Elements<Cell>())
                {
                    foreach (Cell uc in updateFormulaInServiceRow)
                        if (uc.CellReference.Value == cell.CellReference.Value) { UpdateFormula(cell, Offset, CountAddRows, ServiceRow); break; }

                    cell.CellReference = SpreadsheetReader.ColumnFromReference(cell.CellReference) + r.RowIndex.ToString();

                    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    {
                        int itemIndex = int.Parse(cell.CellValue.Text);

                        SharedStringTablePart shareStringTablePart = RecultDoc.WorkbookPart.SharedStringTablePart;
                        if (shareStringTablePart != null)
                        {
                            SharedStringItem item = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(itemIndex);
                            if (item != null) SetFrmulaInServiceCell(item, cell, firstRow, lastRow, countRows);
                        }
                    }
                    if (cell.CellFormula != null)
                    {
                        cell.CellFormula.CalculateCell = true;
                        if (cell.CellValue != null) cell.CellValue.Remove();
                    }

                }
                sheetData.Append(r);
            }

            AppendMergeCells(ListMCellInServiceRow, RecultWorkSheetPart, Offset);
        }

        private static void SetFrmulaInServiceCell(SharedStringItem item, Cell cell, uint firstRow, uint lastRow, int rowCount)
        {
            string newRangeFirst = SpreadsheetReader.ColumnFromReference(cell.CellReference) + firstRow.ToString();
            string newRangeEnd = SpreadsheetReader.ColumnFromReference(cell.CellReference) + lastRow.ToString();
            switch (item.Text.Text.ToLower()) 
            {
                case "sum": cell.CellFormula = new CellFormula("SUM(" + newRangeFirst + ":" + newRangeEnd + ")"); break;
                case "min": cell.CellFormula = new CellFormula("MIN(" + newRangeFirst + ":" + newRangeEnd + ")"); break;
                case "max": cell.CellFormula = new CellFormula("MAX(" + newRangeFirst + ":" + newRangeEnd + ")"); break;
                case "avg": cell.CellFormula = new CellFormula("AVERAGE(" + newRangeFirst + ":" + newRangeEnd + ")"); break;
                case "count": cell.CellFormula = new CellFormula(Convert.ToString(rowCount)); break;
            }

            if (cell.CellFormula != null)
            {
                cell.CellFormula.CalculateCell = true;
                if (cell.CellValue != null) cell.CellValue.Remove();
                if (cell.DataType != null) cell.DataType = CellValues.Number;
            }

            if (lastRow < firstRow)
            {
                if (cell.CellValue != null) cell.CellValue.Remove();
                if (cell.CellFormula != null) cell.CellFormula.Remove();
            }
           
        }


        private static void DeleteRows(WorksheetPart worksheetPart, uint rowIndex, uint count)
        {
                   
          //  WorksheetWriter.DeleteRows(rowIndex, count, worksheetPart);

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

           //Remove rows to delete
            foreach (var row in sheetData.Elements<Row>().
                Where(r => r.RowIndex.Value >= rowIndex && r.RowIndex.Value < rowIndex + count).ToList())
            {
                row.Remove();
            }

           //Get all the rows which are equal or greater than this row index + row count
            IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex.Value >= rowIndex + count);

           //Move the cell references up by the number of rows
            foreach (Row row in rows)
            {
                row.RowIndex.Value -= count;

                IEnumerable<Cell> cells = row.Elements<Cell>();
                foreach (Cell cell in cells)
                {
                    cell.CellReference = SpreadsheetReader.ColumnFromReference(cell.CellReference)
                        + row.RowIndex.Value.ToString();
                }
            }

           MergeCells mergeCells = worksheetPart.Worksheet.GetFirstChild<MergeCells>();
           if (mergeCells != null)
           {

               //Удаление  обединенных ячеек попавших в облать удаления
               foreach (MergeCell mCell in mergeCells.Elements<MergeCell>().
                   Where(mc => SpreadsheetReader.RowFromReference(FirstRefFromRange(mc.Reference.Value)) >= rowIndex 
                       && SpreadsheetReader.RowFromReference(LastRefFromRange(mc.Reference.Value)) <= (rowIndex + count - 1)).ToList())
               {
                   mCell.Remove();
               }


               //Обновления ссылок всех ниже стоящих обединенных ячеек
               foreach (MergeCell mCell in mergeCells.Elements<MergeCell>())
               {
                   string Range = mCell.Reference.Value;
                   string StartRef = FirstRefFromRange(Range);
                   string EndRef = LastRefFromRange(Range);
                   UInt32 startRow = SpreadsheetReader.RowFromReference(StartRef);
                   UInt32 endRow = SpreadsheetReader.RowFromReference(EndRef);

                   if (startRow > (rowIndex + count - 1))
                   {
                       string newRangeFirst = SpreadsheetReader.ColumnFromReference(StartRef) + (SpreadsheetReader.RowFromReference(StartRef) - count);
                       string newRangeEnd = SpreadsheetReader.ColumnFromReference(EndRef) + (SpreadsheetReader.RowFromReference(EndRef) - count);
                       mCell.Reference.Value = newRangeFirst + ":" + newRangeEnd;
                   }
               }
           }
        }

        private static UInt32 AppendRows(DataRow dataRow, IEnumerable<Row> ListRowInRange, IEnumerable<MergeCell> ListMergeCells, List<Cell> updateFormulaInRange, WorksheetPart RecultWorkSheetPart, UInt32 Offset, UInt32 CountAddRow)
        {
            UInt32 CountRangeData = Convert.ToUInt32(ListRowInRange.Count());
            SheetData sheetData = RecultWorkSheetPart.Worksheet.Elements<SheetData>().First();
            foreach (Row RangeRow in ListRowInRange) //Добавляем строки выбраные по указаному DefinedName 
            {
                Row r = (Row)RangeRow.Clone();
                r.RowIndex = r.RowIndex + Offset;
                foreach (Cell cell in r.Elements<Cell>())
                {
                    foreach (Cell uc in updateFormulaInRange)
                    {
                        if (uc.CellReference.Value == cell.CellReference.Value)
                        {
                            UpdateFormula(cell, Offset, CountAddRow, ListRowInRange); break;
                        }
                    }

                    cell.CellReference = SpreadsheetReader.ColumnFromReference(cell.CellReference) + r.RowIndex.ToString();
                    if (dataRow != null)
                    {
                        SetValueToCellFormula(dataRow.Table.DataSet, dataRow, cell);

                        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                        {
                            var SharedString = sharedStringTable.ChildElements[Int32.Parse(cell.CellValue.InnerText)].CloneNode(true);
                            var parcing_Xml = ParcingCellValue(dataRow.Table.DataSet, dataRow, SharedString.InnerXml);
                            if (String.Compare(parcing_Xml, SharedString.InnerXml) > 0)
                            {
                                int i = sharedStringTable.Count();
                                SharedString.InnerXml = parcing_Xml;
                                sharedStringTable.Append(SharedString);

                                cell.CellValue = new CellValue(i.ToString());
                                cell.DataType.Value = CellValues.SharedString;
                                sharedStringTable.Save();
                            }
                        }
                    }

                }
                sheetData.Append(r);
            }

            AppendMergeCells(ListMergeCells, RecultWorkSheetPart, Offset);//Добавление обединенных ячеек в MergeCells присутствующих в ListMergeCells 

            return CountRangeData;
        }

        private static void AppendMergeCells(IEnumerable<MergeCell> ListMergeCells, WorksheetPart RecultWorkSheetPart, uint Offset)
        {
            if (ListMergeCells.Count() > 0)
            {
                foreach (MergeCell mCell in ListMergeCells) //Добавление обединенных ячеек в MergeCells присутствующих в ListMergeCells 
                {
                    string StartRef = FirstRefFromRange(mCell.Reference.Value);
                    string EndRef = LastRefFromRange(mCell.Reference.Value);
                    string newRangeFirst = SpreadsheetReader.ColumnFromReference(StartRef) + (SpreadsheetReader.RowFromReference(StartRef) + Offset);
                    string newRangeEnd = SpreadsheetReader.ColumnFromReference(EndRef) + (SpreadsheetReader.RowFromReference(EndRef) + Offset);
                    MergeCell mergeCell = new MergeCell() { Reference = new StringValue(newRangeFirst + ":" + newRangeEnd) };
                    RecultWorkSheetPart.Worksheet.GetFirstChild<MergeCells>().Append(mergeCell);
                    
                }
            }
        }


        private static void SetValueToCellFormula(DataSet dataSet, DataRow dataRow, Cell cell)
        {
         //   cell.CellFormula.FormulaType = For
            if (cell.CellFormula != null)
                if (cell.CellFormula.Text.Trim().Length > 0)
                {
                    bool recalcFormula = true;
                    string RangeName = "", formula = cell.CellFormula.Text.Trim();

                    if (formula[formula.Length - 1] == '?')
                    {
                        formula = formula.TrimEnd('?');
                        recalcFormula = false;
                    }

                    if (dataRow != null) RangeName = dataRow.Table.TableName;
                    foreach (DataTable Table in dataSet.Tables)
                    {
                        foreach (DataColumn Dcol in Table.Columns)
                        {
                            string val ;
                            string teg = Table.TableName + "_" + Dcol.ColumnName;

                            if (teg.ToLower() == cell.CellFormula.Text.Trim().ToLower())
                            {
                                var cell_value = "";
                                cell.CellValue.Remove();

                                if (dataRow != null)
                                {
                                    cell_value = dataRow[Dcol.ColumnName].ToString();
                                }
                                else
                                {
                                    cell_value = Table.Rows[0][Dcol.ColumnName].ToString();
                                }

                                if ((Dcol.DataType.Name == "Decimal" || Dcol.DataType.Name == "Double" || Dcol.DataType.Name == "Int32" || Dcol.DataType.Name == "Int64"))
                                {
                                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);

                                    if (!String.IsNullOrEmpty(cell_value))
                                    {
                                        cell.CellValue = new CellValue(Convert.ToDecimal(cell_value).ToString("G", CultureInfo.InvariantCulture));
                                    }
                                }
                                else if (Dcol.DataType.Name == "DateTime" )
                                {
                                  /*  cell.DataType = new EnumValue<CellValues>(CellValues.Date);

                                    if (!String.IsNullOrEmpty(cell_value))
                                    {
                                        cell.CellValue = new CellValue(Convert.ToDateTime(cell_value).ToString("s"));
                                    }*/

                                    cell.DataType = null;
                                    if (!String.IsNullOrEmpty(cell_value))
                                    {
                                        cell.CellValue = new CellValue(Convert.ToDateTime(cell_value).ToOADate().ToString("G", CultureInfo.InvariantCulture));
                                    }

                                }
                                else
                                {
                                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    cell.CellValue = new CellValue(cell_value);
                                }

                                cell.CellFormula.Remove();
                                return;
                            }

                            var pattern = @"(\W|^)(" + teg + @")(\W|$)";

                            if (Regex.IsMatch(formula, pattern))
                            {
                                bool str = (Dcol.DataType.Name == "String" || Dcol.DataType.Name == "DateTime");
                                bool Dec = (Dcol.DataType.Name == "Decimal" || Dcol.DataType.Name == "Double");
                                bool Int = (Dcol.DataType.Name == "Int32" || Dcol.DataType.Name == "Int64");

                                bool nullDB;
                                if (RangeName == Table.TableName)
                                {
                                    val = dataRow[Dcol.ColumnName].ToString();
                                    nullDB = Convert.IsDBNull(dataRow[Dcol.ColumnName]);
                                }
                                else
                                {
                                    val = Table.Rows[0][Dcol.ColumnName].ToString();
                                    nullDB = Convert.IsDBNull(Table.Rows[0][Dcol.ColumnName]);
                                }

                                if (str)
                                {
                                    val = "\"" + val.Replace("\"", "\"\"") + "\"";
                                }
                                if (Dec)
                                {
                                    val = val.Replace(',', '.');
                                }

                                if ((Dec || Int) && val.Length == 0)
                                {
                                    val = "0";
                                }

                                formula = Regex.Replace(formula, pattern, "${1}" + val + "${3}"); //formula = formula.Replace(teg, val);

                                if ((Dec || Int) && formula == "0" && nullDB)
                                {
                                    formula = "";
                                }
                            }

                        }
                    }

                    cell.CellFormula.Text = formula;
                    cell.CellFormula.CalculateCell = true;

                    if (!recalcFormula)
                    {
                        ConvertStrFormulaToValue(cell);
                    }
                    else if (cell.CellValue != null) { cell.CellValue.Remove(); }

                }
        }

        private static void ConvertStrFormulaToValue(Cell cell)
        {
            if (cell.CellFormula != null)
            {
                string recult = "";

                string[] split = cell.CellFormula.Text.Split(new Char[] { '&'/*, '\"'*/ });

                foreach (string s in split)
                {
                    recult += s.Trim('\"');
                }
                cell.CellValue.Text = recult;
                cell.DataType.Value = CellValues.String;
                cell.CellFormula.Remove();
            }
        }

        //Первая ссылка в области ссылок
        private static string FirstRefFromRange(string Range)
        {
            return Range.Substring(0, Range.IndexOf(":"));
        }

        //Последняя ссылка в области ссылок
        private static string LastRefFromRange(string Range)
        {
            return Range.Substring(Range.IndexOf(":") + 1);
        }
        //Номер строки соответветствующий началу области
        public static UInt32 FirstRowInRange(string Range)
        {
            return SpreadsheetReader.RowFromReference(FirstRefFromRange(Range));
        }

        //Номер строки соответсвующий концу области
        public static UInt32 LastRowInRange(string Range)
        {
            return SpreadsheetReader.RowFromReference(LastRefFromRange(Range));
        }

        //Получаем имя области (DefinedName)
        public static string GetSheetNameOnDefName(DefinedName d)
        {
            return d.Text.Split('!')[0].Trim('\'');
        }
       
        //Получаем саму область (DefinedName)
        public static string GetRangeInDefName(DefinedName d)
        {
            return d.Text.Split('!')[1];
        }

        //Возвращаем даные поля и строки с идексом  rowIdx соответсвующее тегу,  Teg = (TableName_ColumnName)
        private static string GetDataFromTable(DataSet dataSet, string Teg, int rowIdx)
        {
            string recult = "";
            foreach (DataTable Table in dataSet.Tables)
            {
                foreach (DataColumn Dcol in Table.Columns)
                {
                    if (Teg.ToUpper() == (Table.TableName + "_" + Dcol.ColumnName).ToUpper())
                    {
                        if (Table.Rows.Count > 0)
                        {
                            recult = Table.Rows[rowIdx][Dcol.ColumnName].ToString();

                        }
                    }
                }
            }
            return recult;
        }

        //Ищем DataTable в DataSet за указаными параметрами (TableName)
        private static DataTable GetTableByName(DataSet dataSet, string TableName)
        {
            DataTable rez = null;
            foreach (DataTable Table in dataSet.Tables)
            {
                if (Table.TableName == TableName)
                {
                    rez = Table;
                    break;
                }
            }
            return rez;
        }

        private static void ClearSpreadsheetDocument(SpreadsheetDocument Document)
        {
            Document.WorkbookPart.DeletePart(Document.WorkbookPart.CalculationChainPart);

            foreach (Sheet curSheet in Document.WorkbookPart.Workbook.Sheets)
            {
                WorksheetPart workSheetPart = SpreadsheetReader.GetWorksheetPartByName(Document, curSheet.Name);
                SheetData sheetData = workSheetPart.Worksheet.Elements<SheetData>().First();
                if (!sheetData.HasChildren)
                {
                    MergeCells p = workSheetPart.Worksheet.GetFirstChild<MergeCells>();
                    if (p != null) p.Remove();
                }
            }

        }

        private static void UpdateFormula(Cell cell, UInt32 offset, UInt32 CountAddRow, IEnumerable<Row> ListRowInRange)
        {
            ExcelFormula excelFormula = new ExcelFormula("=" + cell.CellFormula.Text);
            string fStr = "";

            for (int i = 0; i < excelFormula.Count; i++)
            {

                ExcelFormulaToken token = excelFormula[i];
                string tokenValue = token.Value;

                if (token.Subtype == ExcelFormulaTokenSubtype.Start) tokenValue += "(";
                if (token.Subtype == ExcelFormulaTokenSubtype.Text) tokenValue = '\"' + tokenValue.Replace("\"", "\"\"") + '\"';
                if (token.Subtype == ExcelFormulaTokenSubtype.Reference)
                {
                    UInt32 row = SpreadsheetReader.RowFromReference(tokenValue);
                    string col = SpreadsheetReader.ColumnFromReference(tokenValue);
                    if (row > ListRowInRange.ElementAt(ListRowInRange.Count() - 1).RowIndex)
                        tokenValue = col + (row + CountAddRow);
                    else if (row >= ListRowInRange.ElementAt(0).RowIndex) tokenValue = col + (row + offset);
                }
                if (token.Subtype == ExcelFormulaTokenSubtype.Range)
                {
                    string fr = FirstRefFromRange(tokenValue);
                    string lr = LastRefFromRange(tokenValue);
                    if (SpreadsheetReader.RowFromReference(fr) > ListRowInRange.ElementAt(ListRowInRange.Count() - 1).RowIndex)
                        tokenValue = SpreadsheetReader.ColumnFromReference(fr) + (SpreadsheetReader.RowFromReference(fr) + CountAddRow) + ":" + SpreadsheetReader.ColumnFromReference(lr) + (SpreadsheetReader.RowFromReference(lr) + CountAddRow);
                    else if (SpreadsheetReader.RowFromReference(fr) >= ListRowInRange.ElementAt(0).RowIndex)
                        tokenValue = SpreadsheetReader.ColumnFromReference(fr) + (SpreadsheetReader.RowFromReference(fr) + offset) + ":" + SpreadsheetReader.ColumnFromReference(lr) + (SpreadsheetReader.RowFromReference(lr) + offset);
                }
                if (token.Subtype == ExcelFormulaTokenSubtype.Stop) tokenValue = ")";
                fStr += tokenValue;
            }
            cell.CellFormula.Text = fStr;
        
        }

        private static object GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            if (cell.CellValue == null)
            {
                return null;
            }

            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                CellFormat cf = document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[int.Parse(cell.StyleIndex.InnerText)] as CellFormat;

                if (cf.NumberFormatId >= 14 && cf.NumberFormatId <= 22)
                {
                    return DateTime.Parse("1900-01-01").AddDays(int.Parse(cell.CellValue.Text) - 2);
                }
                else
                {
                    return value;
                }
            }
        }
    }
}
