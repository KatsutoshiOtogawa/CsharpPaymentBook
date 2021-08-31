using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace IO
{
    public class Class1
    {
        // 一時ファイルにコピーを受け取って
        static public string OutputTmpSpreadSheet(string TemplateName){
            var excel_template = Environment.GetEnvironmentVariable("EXCEL_TEMPLATE");

            var filepath = Path.Combine(excel_template,TemplateName);

            var tmpfileName = Path.GetTempFileName();

            // バイトをすべて読み込む。
            var docAsArray = File.ReadAllBytes(filepath);

            using (var stream = new MemoryStream())
            {
                stream.Write(docAsArray, 0, docAsArray.Length);
                //作成した一時ファイルにすべて書き込む。
                File.WriteAllBytes(tmpfileName, stream.ToArray());    
            }

            // 作成したファイルをテンプレートから読み込む。
            using (var spreadsheetDoc = SpreadsheetDocument.Open(tmpfileName, true))
            {
                // Change from template type to workbook type
                spreadsheetDoc.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
            }

            return tmpfileName;
        }

        static public void CreateSpreadSheetFromTemplate(string TemplateName,string CreateFileName){

            var tmpfileName = OutputTmpSpreadSheet(TemplateName);

            var excel_output = Environment.GetEnvironmentVariable("EXCEL_OUTPUT");
            
            var outputpath = Path.Combine(excel_output,CreateFileName);

            File.Copy(tmpfileName,outputpath);
        }
    }
}
