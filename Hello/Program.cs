using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
 
namespace Hello
{
    class Program
    {
        static void Main(string[] args)
        {
            String filepath = @"data.xlsx";
 
            // ブック(book)の追加
            var spreadsheetDocument = SpreadsheetDocument.
                Create(filepath, SpreadsheetDocumentType.Workbook);
 
            var workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
 
            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
 
            // シート(sheet)の追加
            var sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());
 
            var sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);
 
            // 行(row)の追加
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var row = new Row() { RowIndex = 1 };
            sheetData.Append(row);
 
            // セル(cell)の追加
            Cell refCell = null;
            var newCell = new Cell()
            {
                CellReference = "A1",
                DataType = CellValues.String,
                CellValue = new CellValue("Hello, Excel")
            };
            row.InsertBefore(newCell, refCell);
 
            // ブックの保存
            workbookpart.Workbook.Save();
 
            // Close the document.
            spreadsheetDocument.Close();
        }
    }
}