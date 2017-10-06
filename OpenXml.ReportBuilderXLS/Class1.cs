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

namespace OpenXml.ReportBuilderXLS
{
    public class ReportBuilderXLS
    {
        int ProgressValue = 0, MaximumCount=0;

        public MemoryStream GenerateReport(Dictionary<string, IList> dictionary, string templateFileName, string OutFile) 
        {
            DataSet dataSet = new DataSet();
            foreach (var obj in dictionary)
            {
                DataTable dataTable = new DataTable(obj.Key);

                object row = dictionary[obj.Key][0];
                Type type = row.GetType();
                PropertyInfo[] pi = type.GetProperties();
                foreach (PropertyInfo info in pi)
                {
                    PropertyInfo p = type.GetProperty(info.Name);
                    dataTable.Columns.Add(info.Name, p.GetValue(row, null).GetType());
                }


                foreach (var obj2 in obj.Value)
                {
                    object personData = obj2;
                    Type personType = personData.GetType();
                    PropertyInfo[] pis = personType.GetProperties();
                    DataRow newRow = dataTable.NewRow();

                    foreach (PropertyInfo pin in pis)
                    {
                        PropertyInfo propertyInfo = personType.GetProperty(pin.Name);
                        newRow[pin.Name] = propertyInfo.GetValue(personData, null);
                    }
                    dataTable.Rows.Add(newRow);
                }
                dataTable.AcceptChanges();
                dataSet.Tables.Add(dataTable);
            }

            return GenerateReport(dataSet, templateFileName, OutFile);
        }

        //Создание отчета
        public MemoryStream GenerateReport(DataSet dataSet, string Temleyt, string OutFile)
        {
            MemoryStream ms = new MemoryStream();
            ms = SpreadsheetReader.Copy(Temleyt);
            SpreadsheetDocument RecultDoc = SpreadsheetDocument.Open(ms, true);

            MemoryStream tp = new MemoryStream();
            tp = SpreadsheetReader.Copy(Temleyt);
            SpreadsheetDocument TemleytDoc = SpreadsheetDocument.Open(tp, true);

            foreach (Sheet curSheet in TemleytDoc.WorkbookPart.Workbook.Sheets)
            {
                UInt32 Offset = 0;
                // Получаем имена областей DefinedNames принадлежащих к текущему Sheet
                List<DefinedName> dn = RecultDoc.WorkbookPart.Workbook.Elements<DefinedNames>().First().Elements<DefinedName>()
                                                             .Where(d => d.Text.IndexOf("#REF!") < 0 && curSheet.Name.Value == d.Text.Substring(0, d.Text.IndexOf("!")).TrimEnd('\'').TrimStart('\'')).ToList();
                foreach (DefinedName definedName in dn)
                {
                    DataTable dataTable = GetTableByName(dataSet, definedName.Name);
                    if (dataTable != null) Offset = SetValueFromDataTable(dataTable, RecultDoc, TemleytDoc, Offset);
                }
           
            }
            
            SetValueWithoutRange(dataSet, RecultDoc);
            ClearSpreadsheetDocument(RecultDoc);

            TemleytDoc.Close();
            RecultDoc.Close();

            if (OutFile != null)
                if (OutFile.Length > 0) SpreadsheetWriter.StreamToFile(OutFile, ms);

            return ms;
        }

        private UInt32 SetValueFromDataTable(DataTable dataTable, SpreadsheetDocument RecultDoc, SpreadsheetDocument TemleytDoc, UInt32 CurOffset)
        {
            UInt32 newOffset = 0;
            string RangeName = dataTable.TableName;
         //   SpreadsheetDocument TemleytDoc = SpreadsheetDocument.Open(Temleyt, true);

            DefinedName DefName = SpreadsheetReader.GetDefinedName(TemleytDoc, RangeName);
            if (DefName != null)
                if (DefName.Text.IndexOf("#REF!") < 0)
                {
                    string definedName = DefName.Text;
                    string WorkSheetName = definedName.Substring(0, definedName.IndexOf("!")).TrimEnd('\'').TrimStart('\'');
                    string Range = definedName.Substring(definedName.IndexOf("!") + 1);
                    UInt32 startRowInRange = FirstRowInRange(Range);
                    UInt32 endRowInRange = LastRowInRange(Range);
                    UInt32 serviceRow = 1;
                    UInt32 CountRangeData = (endRowInRange - startRowInRange + 1) - serviceRow;

                    WorksheetPart TemleytWorkSheet = SpreadsheetReader.GetWorksheetPartByName(TemleytDoc, WorkSheetName);

                    List<Row> ListRowInRange = TemleytWorkSheet.Worksheet.Elements<SheetData>().First().Elements<Row>()
                                                                      .Where(r => r.RowIndex >= startRowInRange && r.RowIndex < endRowInRange).ToList();
                    List<Row> ListRowBottom = TemleytWorkSheet.Worksheet.Elements<SheetData>().First().Elements<Row>()
                                                                      .Where(r => r.RowIndex > endRowInRange).ToList();
                    List<Row> ServiceRow = TemleytWorkSheet.Worksheet.Elements<SheetData>().First().Elements<Row>()
                                                                      .Where(r => r.RowIndex == endRowInRange).ToList();

                    List<MergeCell> ListMCellInRange = new List<MergeCell>() ;
                    List<MergeCell> ListMCellBottom = new List<MergeCell>();
                    MergeCells mergeCells = TemleytWorkSheet.Worksheet.GetFirstChild<MergeCells>();
                    if (mergeCells!=null)
                    {
                        ListMCellInRange = TemleytWorkSheet.Worksheet.Elements<MergeCells>().First().Elements<MergeCell>()
                                                                    .Where(r => SpreadsheetReader.RowFromReference(FirstRefFromRange(r.Reference.Value)) >= startRowInRange && SpreadsheetReader.RowFromReference(LastRefFromRange(r.Reference.Value)) <= endRowInRange).ToList();
                        ListMCellBottom = TemleytWorkSheet.Worksheet.Elements<MergeCells>().First().Elements<MergeCell>()
                                                                    .Where(r => SpreadsheetReader.RowFromReference(FirstRefFromRange(r.Reference.Value)) > endRowInRange).ToList();
                    }
                    UInt32 NewStartRow = startRowInRange + CurOffset;

                    WorksheetPart RecultWorkSheetPart = SpreadsheetReader.GetWorksheetPartByName(RecultDoc, WorkSheetName);
                    WorksheetWriter RecultWriter = new WorksheetWriter(RecultDoc, RecultWorkSheetPart);
                    DeleteRows(RecultDoc, RecultWriter, NewStartRow, CountRangeData + serviceRow);
                    if (ListRowBottom.Count() > 0) DeleteRows(RecultDoc, RecultWriter, NewStartRow, ListRowBottom.Last().RowIndex.Value + CurOffset - NewStartRow + 1);

                    MaximumCount = dataTable.Rows.Count;
                  
                    UInt32 InsideOffset = NewStartRow;
                    foreach (DataRow dataRow in dataTable.Rows)
                    {
                        UInt32 Offset = (InsideOffset - NewStartRow) + CurOffset;
                        if (dataRow.RowState.ToString() != "Deleted")
                        {
                            AppendRows(dataRow, ListRowInRange, ListMCellInRange, RecultWorkSheetPart, Offset);
                            InsideOffset += CountRangeData;
                        }
                        ++ProgressValue;
                    }

                    newOffset = (InsideOffset - NewStartRow - CountRangeData) + CurOffset;
                    AppendRows(null, ListRowBottom, ListMCellBottom, RecultWorkSheetPart, newOffset);

                    DefName = SpreadsheetReader.GetDefinedName(RecultDoc, RangeName);
                    definedName = DefName.Text;
                    DefName.Text = definedName.Substring(0, definedName.LastIndexOf('$') + 1) + InsideOffset.ToString();

                 //   DefName.Text = definedName.Substring(0, definedName.IndexOf("!") + 1) + '$' + SpreadsheetReader.ColumnFromReference(FirstRefFromRange(Range)) + '$' + NewStartRow.ToString() + ":$" + SpreadsheetReader.ColumnFromReference(LastRefFromRange(Range)) + '$' + StartClone.ToString();

                    //Обновляем ссилки начала и конца рабочих областей DefinedNames
                    foreach (DefinedName CurdefName in RecultDoc.WorkbookPart.Workbook.DefinedNames)
                    {
                        if (CurdefName.Text.IndexOf("#REF!") < 0)
                        {
                         //   if (CurdefName.Text == definedName)//Текущая облать DefinedName
                        //    {
                       //         CurdefName.Text = definedName.Substring(0, definedName.IndexOf("!") + 1) + '$' + SpreadsheetReader.ColumnFromReference(FirstRefFromRange(Range)) + '$' + NewStartRow.ToString() + ":$" + SpreadsheetReader.ColumnFromReference(LastRefFromRange(Range)) + '$' + StartOffset.ToString();
                      //      }
                       //     else //Все остальные DefinedName текущего листа Sheet
                      //      {
                                string strDefName = CurdefName.Text;
                                string sheetName = strDefName.Substring(0, strDefName.IndexOf("!")).Trim('\'');
                                string range = strDefName.Substring(strDefName.IndexOf("!") + 1);
                                if (sheetName == WorkSheetName)
                                {
                                    string firstRefInRange = FirstRefFromRange(range);
                                    string lastRefInRange = LastRefFromRange(range);
                                    UInt32 firstRow = SpreadsheetReader.RowFromReference(firstRefInRange);
                                    UInt32 lastRow = SpreadsheetReader.RowFromReference(lastRefInRange);
                                    if (firstRow + newOffset > InsideOffset + CurOffset)
                                    {
                                        CurdefName.Text = strDefName.Substring(0, strDefName.IndexOf("!") + 1) + '$' + SpreadsheetReader.ColumnFromReference(firstRefInRange) + '$' + (firstRow + newOffset).ToString() + ":$" + SpreadsheetReader.ColumnFromReference(lastRefInRange) + '$' + (lastRow + newOffset).ToString();
                                    }
                                }

                  //          }
                        }
                    }

                }
         //   TemleytDoc.Close();
            return newOffset;
        }

        private void DeleteRows(SpreadsheetDocument RecultDoc, WorksheetWriter RecultWriter, uint rowIndex, uint count)
        {
            RecultWriter.DeleteRows(rowIndex, count);

            if (RecultWriter.Worksheet.Worksheet.Elements<MergeCells>().Count() > 0)
            {
                MergeCells mergeCells = RecultWriter.Worksheet.Worksheet.Elements<MergeCells>().First();

                //Удаление  обединенных ячеек попавших в облать удаления
                for (int i = mergeCells.Elements<MergeCell>().Count() - 1; i >= 0; i--)
                {
                    string Range = mergeCells.Elements<MergeCell>().ElementAt(i).Reference.Value;
                    UInt32 startRow = SpreadsheetReader.RowFromReference(FirstRefFromRange(Range));
                    UInt32 endRow = SpreadsheetReader.RowFromReference(LastRefFromRange(Range));

                    if (startRow >= rowIndex && endRow <= (rowIndex + count - 1))
                    {
                        mergeCells.Elements<MergeCell>().ElementAt(i).Remove();
                    }
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

        private UInt32 AppendRows(DataRow dataRow, IEnumerable<Row> ListRowInRange, IEnumerable<MergeCell> ListMergeCells, WorksheetPart RecultWorkSheetPart, UInt32 Offset)
        {
            // String RangeName = "";
            //  if (dataRow != null) RangeName = dataRow.Table.TableName;
            UInt32 CountRangeData = Convert.ToUInt32(ListRowInRange.Count());
            SheetData sheetData = RecultWorkSheetPart.Worksheet.Elements<SheetData>().First();
            foreach (Row RangeRow in ListRowInRange) //Добавляем строки выбраные по указаному DefinedName 
            {
                Row r = (Row)RangeRow.Clone();
                r.RowIndex = r.RowIndex + Offset;
                foreach (Cell cell in r.Elements<Cell>())
                {
                    cell.CellReference = SpreadsheetReader.ColumnFromReference(cell.CellReference) + r.RowIndex.ToString();
                    if (dataRow != null)
                    {
                        // ConvertMathFormulaToCellValue(dataRow.Table.DataSet, dataRow.Table.Columns, dataRow, cell);
                        // ConvertFormulaToValue(dataRow.Table.DataSet, dataRow.Table.Columns, dataRow, cell);
                        SetValueToCellFormula(dataRow.Table.DataSet, dataRow, cell);
                    }
                }
                sheetData.Append(r);
            }

            if (ListMergeCells.Count() > 0)
            {
                foreach (MergeCell mCell in ListMergeCells) //Добавление обединенных ячеек присутствующих в ListMergeCells 
                {
                    string StartRef = FirstRefFromRange(mCell.Reference.Value);
                    string EndRef = LastRefFromRange(mCell.Reference.Value);
                    string newRangeFirst = SpreadsheetReader.ColumnFromReference(StartRef) + (SpreadsheetReader.RowFromReference(StartRef) + Offset);
                    string newRangeEnd = SpreadsheetReader.ColumnFromReference(EndRef) + (SpreadsheetReader.RowFromReference(EndRef) + Offset);
                    MergeCell mergeCell = new MergeCell() { Reference = new StringValue(newRangeFirst + ":" + newRangeEnd) };
                    RecultWorkSheetPart.Worksheet.GetFirstChild<MergeCells>().Append(mergeCell);
                }

            }
            return CountRangeData;
        }

        class FormulaEntry
        {
            int point;
            string str;
        }

        private void SetValueToCellFormula(DataSet dataSet, DataRow dataRow, Cell cell)
        {
            
            if (cell.CellFormula != null)
            {
                List<FormulaEntry> Formula = new List<FormulaEntry>();
                string RangeName = "", formula = cell.CellFormula.Text;
                if (dataRow != null) RangeName = dataRow.Table.TableName;
                foreach (DataTable Table in dataSet.Tables)
                {
                    foreach (DataColumn Dcol in Table.Columns)
                    {
                        string val;
                        bool str = (Dcol.DataType.Name == "String" || Dcol.DataType.Name == "DateTime");
                        bool Dec = (Dcol.DataType.Name == "Decimal" || Dcol.DataType.Name == "Double" );
                        if (RangeName == Table.TableName) val = dataRow[Dcol.ColumnName].ToString();
                        else val = Table.Rows[0][Dcol.ColumnName].ToString();

                        if (str) val = "\"" + val.Replace("\"", "\"\"") + "\"";
                        if (Dec) val = val.Replace(',', '.');
                        if (Dcol.DataType.Name == "Int32" && val.Length == 0) val = "0";

                        string teg = Table.TableName + "_" + Dcol.ColumnName;
                        int first,last;
                        first = formula.IndexOf(teg);
                        last = first+teg.Length;

                        //IfSumbol(char firstchar, char lastchar)
                        formula = formula.Replace(teg, val);
                    }
                }

                cell.CellFormula.Text = formula; // cell.CellFormula.CalculateCell = true;
                if (cell.CellValue != null) cell.CellValue.Remove();
            }
        }

        private Cell ConvertStrFormulaToValue(DataSet dataSet, DataColumnCollection Columns, DataRow dataRow, Cell cell)
        {
            if (cell.CellFormula != null)
            {
                string recult = "";

                string[] split = cell.CellFormula.Text.Split(new Char[] { '&', '\"' });

                foreach (string s in split)
                {
                    if (s.Trim().Length > 0)
                    {
                        bool ifColumns = false;
                        if (Columns != null)
                        {
                            string RangeName = dataRow.Table.TableName;
                            foreach (DataColumn Dcol in Columns)
                            {
                                if (s.ToUpper() == (RangeName + "_" + Dcol.ColumnName).ToUpper())
                                {
                                    recult += dataRow[Dcol.ColumnName].ToString();
                                    ifColumns = true;
                                    break;
                                }
                            }
                        }
                        if (!ifColumns && dataSet != null)
                        {
                            string r = GetDataFromTable(dataSet, s, 0);
                            recult += r;
                            ifColumns = (r.Length > 0);
                        }
                        if (!ifColumns) recult += s;
                    }
                }
                cell.CellValue.Text = recult;
                cell.DataType.Value = CellValues.String;
                cell.CellFormula.Remove();

            }
            return cell;
        }

        private void ConvertMathFormulaToCellValue(Cell cell)
        {
            if (cell.CellFormula != null)
            {
                string formula = cell.CellFormula.Text;
                if (formula.LastIndexOfAny(new Char[] { '+', '-', '/', '*' }) < 0) return;


                cell.CellValue.Text = formula;
                cell.DataType.Value = CellValues.String;
                cell.CellFormula.Remove();

            }
        }

        //Первая ссылка в области ссылок
        private string FirstRefFromRange(string Range)
        {
            return Range.Substring(0, Range.IndexOf(":"));
        }

        //Последняя ссылка в области ссылок
        private string LastRefFromRange(string Range)
        {
            return Range.Substring(Range.IndexOf(":") + 1);
        }
        //Номер строки соответветствующий началу области
        private UInt32 FirstRowInRange(string Range)
        {
            return SpreadsheetReader.RowFromReference(FirstRefFromRange(Range));
        }

        //Номер строки соответсвующий концу области
        private UInt32 LastRowInRange(string Range)
        {
            return SpreadsheetReader.RowFromReference(LastRefFromRange(Range));
        }

        //Заполняем все ячейки по всему документу которые непринадлежат ни одному из DefinedName
        private void SetValueWithoutRange(DataSet dataSet, SpreadsheetDocument RecultDoc)
        {
            foreach (Sheet curSheet in RecultDoc.WorkbookPart.Workbook.Sheets)
            {
                WorksheetPart TemleytWorkSheet = SpreadsheetReader.GetWorksheetPartByName(RecultDoc, curSheet.Name);
                SheetData sheetData = TemleytWorkSheet.Worksheet.Elements<SheetData>().First();

                // Получаем имена областей DefinedNames принадлежащих к текущему Sheet
                List<DefinedName> dn = RecultDoc.WorkbookPart.Workbook.Elements<DefinedNames>().First().Elements<DefinedName>()
                                                             .Where(d => d.Text.IndexOf("#REF!") < 0 && curSheet.Name.Value == d.Text.Substring(0, d.Text.IndexOf("!")).TrimEnd('\'').TrimStart('\'')).ToList();

                //Выбираем все строки с SheetData которые не принадлежат ни одному из DefinedName
                List<Row> selctRow = new List<Row>(TemleytWorkSheet.Worksheet.Elements<SheetData>().First().Elements<Row>());
                foreach (DefinedName DN in dn)
                {
                    string Range = DN.Text.Substring(DN.Text.IndexOf("!") + 1);
                    selctRow = selctRow.Where(r => r.RowIndex < FirstRowInRange(Range) || r.RowIndex > LastRowInRange(Range)).ToList();
                }

                //Заполняем выбраные ячейки selctRow даными с dataSet согласно формулам
                foreach (Row curRow in selctRow)
                {
                    foreach (Cell curCell in curRow.Elements<Cell>())
                    {
                        //ConvertFormulaToValue(dataSet, null, null, curCell);
                        SetValueToCellFormula(dataSet, null, curCell);
                    }
                }
            }

        }

        //Возвращаем даные поля и строки с идексом  rowIdx соответсвующее тегу,  Teg = (TableName_ColumnName)
        private string GetDataFromTable(DataSet dataSet, string Teg, int rowIdx)
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
        private DataTable GetTableByName(DataSet dataSet, string TableName)
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

        private void ClearSpreadsheetDocument(SpreadsheetDocument Document)
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

        private bool IfSumbol(char firstchar, char lastchar)
        {
            const string firstcharStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
            const string lastcharStr = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
            return (firstcharStr.IndexOf(firstchar) < 0 && lastcharStr.IndexOf(lastchar) < 0);
        }

    }
}
