using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using IO;

namespace Hello
{
    class Program
    {
        static void Main(string[] args)
        {

            Class1.CreateSpreadSheetFromTemplate(@"tf10378474_win32.xltx",@"tf10378474_win32.xlsx");

            var excel_output = Environment.GetEnvironmentVariable("EXCEL_OUTPUT");
            
            var outputpath = Path.Combine(excel_output,@"tf10378474_win32.xlsx");

            var cellReference = "J4";

            using (var spreadsheetDoc = SpreadsheetDocument.Open(outputpath, true))
            {
                // Change from template type to workbook type
                Sheets sheets = spreadsheetDoc.WorkbookPart.Workbook.Sheets;
                Sheet  sheet  = sheets.Elements<Sheet>().Where(s => s.Name == "請求書2").FirstOrDefault();
                if (null != sheet)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDoc.WorkbookPart.GetPartById(sheet.Id);
                    Worksheet     worksheet     = worksheetPart.Worksheet;
                    Cell          cell          = worksheet.Descendants<Cell>().Where(c => c.CellReference == cellReference).FirstOrDefault();
                    if (null != cell)
                    {
                        string text = cell.InnerText;
                        switch (cell.DataType.Value)
                        {
                            case CellValues.SharedString:
                                int index = int.Parse(cell.InnerText);
                                SharedStringTablePart ssTablePart = spreadsheetDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                text = ssTablePart.SharedStringTable.ElementAt(index).InnerText;
                                break;
 
                            case CellValues.Boolean:
                                if ("0" == cell.InnerText)
                                    text = "FALSE";
                                else
                                    text = "TRUE";
                                break;
                        }
                        System.Console.WriteLine(text);
                    }
                }
            }

            // ブック(book)の追加
            // var spreadsheetDocument = SpreadsheetDocument.
            //     Create(filepath, SpreadsheetDocumentType.Workbook);
 
            // var workbookpart = spreadsheetDocument.AddWorkbookPart();
            // workbookpart.Workbook = new Workbook();
 
            // var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            // worksheetPart.Worksheet = new Worksheet(new SheetData());
 
            // // シート(sheet)の追加
            // var sheets = spreadsheetDocument.WorkbookPart.Workbook.
            //     AppendChild<Sheets>(new Sheets());
 
            // var sheet = new Sheet()
            // {
            //     Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
            //     SheetId = 1,
            //     Name = "mySheet"
            // };
            // sheets.Append(sheet);
 
            // // 行(row)の追加
            // var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            // var row = new Row() { RowIndex = 1 };
            // sheetData.Append(row);
 
            // // セル(cell)の追加
            // Cell refCell = null;
            // var newCell = new Cell()
            // {
            //     CellReference = "A1",
            //     DataType = CellValues.String,
            //     CellValue = new CellValue("Hello, Excel")
            // };
            // row.InsertBefore(newCell, refCell);
 
            // // ブックの保存
            // workbookpart.Workbook.Save();
 
            // // Close the document.
            // spreadsheetDocument.Close();
        }

    }
}