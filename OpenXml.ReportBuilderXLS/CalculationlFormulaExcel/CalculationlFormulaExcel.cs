using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser;
using NCalc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Extensions;
using System.IO;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml;
using System.Globalization;

namespace FormulaExcel
{
    public static class CalculationlFormulaExcel
    {
        public static ProgressBar progressBar;

        public static MemoryStream CalcSpreadsheetDocument(MemoryStream ms, bool DeleteFormula)
        {
            return CalcSpreadsheetDocument(ms, DeleteFormula, "");
        }

        public static MemoryStream CalcSpreadsheetDocument(MemoryStream ms, bool DeleteFormula, string OutFile)
        {
            bool pg = (progressBar != null);
            SpreadsheetDocument RecultDoc = SpreadsheetDocument.Open(ms, true);
            if (DeleteFormula) RecultDoc.WorkbookPart.DeletePart(RecultDoc.WorkbookPart.CalculationChainPart);

            foreach (Sheet curSheet in RecultDoc.WorkbookPart.Workbook.Sheets)
            {
                WorksheetPart RecultWorkSheetPart = SpreadsheetReader.GetWorksheetPartByName(RecultDoc, curSheet.Name.Value);

                IEnumerable<Row> ListRowInRange = RecultWorkSheetPart.Worksheet.Elements<SheetData>().First().Elements<Row>().Where(r => r.Elements<Cell>().Where(c => c.CellFormula != null).Count() > 0);
                if (pg) progressBar.Maximum = ListRowInRange.Count();
                int value = 0;
                foreach (Row Row in ListRowInRange)
                {
                    IEnumerable<Cell> fCell = Row.Elements<Cell>().Where(c => c.CellFormula != null).Where(c => c.CellFormula.Text.Trim() != "");
                    foreach (Cell c in fCell)
                    {
                      var dd =  calcFormule(c, RecultWorkSheetPart, RecultDoc, DeleteFormula);
                    }
                    if (pg) progressBar.Value = ++value;
                }

            }
            RecultDoc.Close();

            if (OutFile != null)
                if (OutFile.Length > 0) SpreadsheetWriter.StreamToFile(OutFile, ms);

            return ms;
        }

        public static string calcFormule(Cell cell, WorksheetPart workSheetPart, SpreadsheetDocument RecultDoc, bool DelCellFormula)
        {
            if (cell.CellValue == null) cell.Append(new CellValue { Text = "" });

            ExcelFormula excelFormula = new ExcelFormula("=" + cell.CellFormula.Text); //Разбираем формулу на узлы
            ReplaceRefInFormulaToValue(excelFormula, workSheetPart, RecultDoc); //Замена ссилок в формулах на их значения
            CompressFormulaToken(excelFormula);  // Сумируем все строковые виражения разделенные знаком & в ExcelFormula

            //Собираем формулу в строковое выражение для расчета в NCalc
            string fStr = "";
            for (int i = 0; i < excelFormula.Count; i++)
            {

                ExcelFormulaToken token = excelFormula[i];
                string tokenValue = token.Value;

                if (token.Subtype == ExcelFormulaTokenSubtype.Start) tokenValue += "(";
                if (token.Subtype == ExcelFormulaTokenSubtype.Text) tokenValue = "'" + tokenValue.Replace(@"\", "/") + "'";
                if (token.Subtype == ExcelFormulaTokenSubtype.Concatenation && token.Value == "&") tokenValue = "+";
                if (token.Subtype == ExcelFormulaTokenSubtype.Number) tokenValue = tokenValue.Replace(',', '.');
                if (token.Subtype == ExcelFormulaTokenSubtype.Stop) tokenValue = ")";
                fStr += tokenValue;
            }

            if (excelFormula.Count == 1) // Если в формуле только один узел заменяем значение ячейки значением формулы
            {
                if (excelFormula[0].Subtype == ExcelFormulaTokenSubtype.Text)
                {

                    cell.CellValue.Text = excelFormula[0].Value.Trim('\'');
                    if (cell.DataType != null) cell.DataType.Value = CellValues.String;
                    if (DelCellFormula) cell.CellFormula.Remove();
                    else cell.CellFormula.CalculateCell = false;
                }
                if (excelFormula[0].Subtype == ExcelFormulaTokenSubtype.Number)
                {
                    cell.CellValue.Text = excelFormula[0].Value;
                    if (cell.DataType != null) cell.DataType.Value = CellValues.Number;
                    if (DelCellFormula) cell.CellFormula.Remove();
                    else cell.CellFormula.CalculateCell = false;
                }
                return cell.CellValue.Text;
            }

            Expression Ev = new Expression(fStr, EvaluateOptions.IgnoreCase | EvaluateOptions.NoCache);
            Ev.EvaluateFunction += EvaluateFunction1;

            try
            {
                
                cell.CellValue.Text = Convert.ToString(Ev.Evaluate());
            //   if (cell.DataType != null)
            //    {
                    decimal r;
                    if (decimal.TryParse(cell.CellValue.Text, out r))
                    {
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        cell.CellValue.Text = r.ToString("G", CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    }
            //    }

                if (DelCellFormula)
                {
                    cell.CellFormula.Remove();
                }
                else
                {
                    cell.CellFormula.CalculateCell = true;
                }
            }
            catch (System.Exception ex)
            {
             //   cell.DataType.Value = CellValues.String;
              //  cell.CellValue.Text = ex.Message + "[" + fStr + "]";
            //    AddComments(workSheetPart, cell.CellReference.Value, "Qw" /*ex.Message + "[" + fStr + "]"*/);
            }
            finally
            {
                ;
            }
            return cell.CellValue.Text;
        }

        private static void AddComments(WorksheetPart workSheetPart, string Ref, string CommentsText)
        {
            VmlDrawingPart vpart;
            if (workSheetPart.VmlDrawingParts == null)
            {
                vpart = workSheetPart.AddNewPart<VmlDrawingPart>("Rid13");
 
            }
          
      //      foreach (VmlDrawingPart vd in workSheetPart.VmlDrawingParts)
     //       {
   //             ;
   //         }
   
       // //    LegacyDrawing lDrawing = new LegacyDrawing();
        //    lDrawing.Id = "Rid13";
        //    workSheetPart.Worksheet.Append(lDrawing);
          //  workSheetPart.Worksheet.Save();
           

            WorksheetCommentsPart commentsPart;
            if (workSheetPart.WorksheetCommentsPart != null) commentsPart = workSheetPart.WorksheetCommentsPart;
            else commentsPart = workSheetPart.AddNewPart<WorksheetCommentsPart>();

            if (commentsPart.Comments == null) commentsPart.Comments = new Comments();
            if (commentsPart.Comments.Authors == null) commentsPart.Comments.Authors = new Authors();

            int authorId = -1;
            foreach (Author at in commentsPart.Comments.Authors)
            {
                authorId++;
                if (at.Text == "ReportBuilder") break;
            }

            if (authorId == -1)
            {
                commentsPart.Comments.Authors.AppendChild<Author>(new Author("ReportBuilder"));
                authorId++;
            }

            if (commentsPart.Comments.CommentList == null) commentsPart.Comments.CommentList = new CommentList();

 
            Comment comment = new Comment();
            comment.Reference = Ref;
            comment.AuthorId = Convert.ToUInt32(authorId);

            comment.CommentText = new CommentText(new Run(new Text(CommentsText)));
            


            commentsPart.Comments.CommentList.AppendChild<Comment>(comment);
            commentsPart.Comments.Save();
        }



        private static void EvaluateFunction1(string name, FunctionArgs args)
        {
            decimal RecultD = 0;
            string RecultStr = "";
            switch (name)
            {
                case "sum":
                    RecultD = 0;
                    for (int a = 0; a < args.Parameters.Length; a++)
                        RecultD += Convert.ToDecimal(args.Parameters[a].Evaluate());
                    args.Result = RecultD;
                    break;

                case "min":
                    RecultD = Convert.ToDecimal(args.Parameters[0].Evaluate());
                    for (int a = 0; a < args.Parameters.Length; a++)
                    {
                        RecultD = Math.Min(RecultD, Convert.ToDecimal(args.Parameters[a].Evaluate()));
                    }
                    args.Result = RecultD;
                    break;

                case "max":
                    RecultD = Convert.ToDecimal(args.Parameters[0].Evaluate());
                    for (int a = 0; a < args.Parameters.Length; a++)
                    {
                        RecultD = Math.Max(RecultD, Convert.ToDecimal(args.Parameters[a].Evaluate()));
                    }

                    args.Result = RecultD;
                    break;

                case "average":
                    for (int a = 0; a < args.Parameters.Length; a++)
                        RecultD += Convert.ToDecimal(args.Parameters[a].Evaluate());

                    args.Result = RecultD / args.Parameters.Length;
                    break;

                case "concatenate":
                    for (int a = 0; a < args.Parameters.Length; a++)
                        RecultStr += Convert.ToString(args.Parameters[a].Evaluate());
                    args.Result = RecultStr;
                    break;
            }

        }

        public static Cell FindCell(string reference, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string columnName = SpreadsheetReader.ColumnFromReference(reference);
            uint rowIndex = SpreadsheetReader.RowFromReference(reference);
            string cellReference = (columnName + rowIndex.ToString());
            int columnIndex = SpreadsheetReader.GetColumnIndex(columnName);

            Row row = null;
            var match = sheetData.Elements<Row>().Where(r => r.RowIndex.Value == rowIndex);

            if (match.Count() != 0)
            {
                row = match.First();
                IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference);
                if ((cells.Count() > 0))
                {
                    return cells.First();
                }
                else return null;
            }
            else return null;
        }

        public static List<Cell> FindCells(string Range, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            Range = Range.Replace("$", "");
            string FirstReference = Range.Split(':')[0], LastReference = Range.Split(':')[1];


            uint FirstrowIndex = SpreadsheetReader.RowFromReference(FirstReference);

            uint LastrowIndex = SpreadsheetReader.RowFromReference(LastReference);

            int FirstcolumnIndex = SpreadsheetReader.GetColumnIndex(SpreadsheetReader.ColumnFromReference(FirstReference));
            int LastcolumnIndex = SpreadsheetReader.GetColumnIndex(SpreadsheetReader.ColumnFromReference(LastReference));

            IEnumerable<Row> match = sheetData.Elements<Row>().Where(r => r.RowIndex.Value >= FirstrowIndex && r.RowIndex.Value <= LastrowIndex);

            List<Cell> listCells = new List<Cell>();
            foreach (Row row in match)
            {
                foreach (Cell cell in row)
                {
                    if (SpreadsheetReader.GetColumnIndex(SpreadsheetReader.ColumnFromReference(cell.CellReference.Value)) >= FirstcolumnIndex && SpreadsheetReader.GetColumnIndex(SpreadsheetReader.ColumnFromReference(cell.CellReference.Value)) <= LastcolumnIndex)
                    {
                        if (cell.CellValue != null)
                        {
                            double d;
                            if (cell.DataType != null)
                            {
                                if (cell.DataType.Value == CellValues.Number || cell.DataType.Value == CellValues.SharedString) listCells.Add(cell);
                                if (cell.DataType.Value == CellValues.String)
                                {
                                    if (double.TryParse(cell.CellValue.Text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), out d))
                                    {
                                        listCells.Add(cell);
                                    }
                                }
                            }
                            else if (double.TryParse(cell.CellValue.Text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), out d))
                            {
                                listCells.Add(cell);
                            }
                        }
                        else if (cell.CellFormula != null) listCells.Add(cell);

                    }
                }
            }

            return listCells;
        }

        private static void ReplaceRefInFormulaToValue(ExcelFormula excelFormula, WorksheetPart workSheetPart, SpreadsheetDocument RecultDoc)
        {
            SharedStringTablePart shareStringTablePart = RecultDoc.WorkbookPart.SharedStringTablePart;
            for (int i = 0; i < excelFormula.Count; i++)
            {
                ExcelFormulaToken token = excelFormula[i];
                if (token.Subtype == ExcelFormulaTokenSubtype.Reference)
                {
                    try
                    {
                        Cell fCell = FindCell(token.Value, workSheetPart);

                        if (fCell.CellValue != null)
                        {
                            if (fCell.DataType.Value == CellValues.SharedString)
                            {
                                int itemIndex = int.Parse(fCell.CellValue.Text);

                                if (shareStringTablePart != null)
                                {
                                    SharedStringItem item = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(itemIndex);
                                    if (item != null) token.Value = item.Text.Text;
                                }
                            }
                            else token.Value = fCell.CellValue.Text;
                        }
                        else if (fCell.CellFormula != null) token.Value = calcFormule(fCell, workSheetPart, RecultDoc, false);
                        else token.Value = "";

                        double d;
                        if (double.TryParse(token.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), out d))
                            token.Subtype = ExcelFormulaTokenSubtype.Number;
                        else token.Subtype = ExcelFormulaTokenSubtype.Text;
                    }
                    catch (System.Exception ex)
                    {
                        ;
                    }
                    finally
                    {
                        ;
                    }

                }
                if (token.Subtype == ExcelFormulaTokenSubtype.Range)
                {
                    try
                    {
                        string str = "";
                        foreach (Cell c in FindCells(token.Value, workSheetPart))
                        {
                            if (c.CellValue != null)
                            {
                                if (c.DataType.Value == CellValues.SharedString)
                                {
                                    int itemIndex = int.Parse(c.CellValue.Text);
                                    if (shareStringTablePart != null)
                                    {
                                        SharedStringItem item = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(itemIndex);
                                        if (item != null) str += item.Text.Text + ",";
                                    }
                                }
                                else str += c.CellValue.Text + ",";
                            }
                            else if (c.CellFormula != null) str += calcFormule(c, workSheetPart, RecultDoc, false) + ",";
                        }
                        token.Value = str.TrimEnd(',');
                    }
                    catch (System.Exception ex)
                    {
                        ;
                    }
                    finally
                    {
                        ;
                    }

                }

            }
        }

        // Сумируем все строковые виражения разделенные знаком & в ExcelFormula
        private static void CompressFormulaToken(ExcelFormula excelFormula)
        {
            while (excelFormula.Where<ExcelFormulaToken>(t => t.Subtype == ExcelFormulaTokenSubtype.Concatenation && t.Value == "&").Count() > 0)
            {
                for (int i = 0; i < excelFormula.Count; i++)
                {
                    ExcelFormulaToken token = excelFormula[i];
                    if (token.Value == "&")
                    {
                        string newTokenValue, firsOperand = "", lastOperand;

                        if (i > 1)
                        {
                            if (excelFormula[i - 2].Type == ExcelFormulaTokenType.OperatorPrefix)
                                firsOperand = excelFormula[i - 2].Value + excelFormula[i - 1].Value;
                            else firsOperand = excelFormula[i - 1].Value;
                        }
                        else firsOperand = excelFormula[i - 1].Value;

                        if (excelFormula[i + 1].Type == ExcelFormulaTokenType.OperatorPrefix)
                        {
                            lastOperand = excelFormula[i + 1].Value + excelFormula[i + 2].Value;
                        }
                        else lastOperand = excelFormula[i + 1].Value;

                        newTokenValue = firsOperand + lastOperand;

                        if (excelFormula[i + 1].Type == ExcelFormulaTokenType.OperatorPrefix)
                        {
                            excelFormula.RemoveAt(i + 1);
                            excelFormula.RemoveAt(i + 1);
                        }
                        else excelFormula.RemoveAt(i + 1);

                        if (i > 1)
                            if (excelFormula[i - 2].Type == ExcelFormulaTokenType.OperatorPrefix)
                            {
                                excelFormula.RemoveAt(i - 1);
                                i--;
                            }

                        excelFormula.RemoveAt(i - 1);
                        i--;
                        excelFormula[i].Value = newTokenValue;
                        excelFormula[i].Subtype = ExcelFormulaTokenSubtype.Text;
                        excelFormula[i].Type = ExcelFormulaTokenType.Operand;

                        break;
                    }
                }
            }
        }

    }

}
