using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficeLib
{
    public class ExcelDocument : IDisposable
    {
        Excel.Application?  app = null;
        //Excel.Workbooks?  books = null;
        Excel.Workbook?    book = null;
        Excel.Sheets?    sheets = null;
        Excel.Worksheet?  sheet = null;
       // Excel.Range?      range = null;

        public ExcelDocument()
        {
            app = new Excel.Application();
            //books = app.Workbooks;
            //book = books.Add();
            book = app.Workbooks.Add();
            sheets = book.Sheets;
            sheet = sheets[1];

        }

        public ExcelDocument(string FileName)
        {
            app = new Excel.Application();
            //books = app.Workbooks;
            //book = books.Open(FileName);
            book = app.Workbooks.Open(FileName); 
            sheets = book.Sheets;
            sheet = sheets[1];
        }

        public void SaveAs(string FileName) 
        { 
            book.SaveAs(FileName);
        }



        public void Dispose()
        {
            book?.Close();
            app?.Quit();

            //Release(range); range = null;
            Release(sheet); sheet = null;
            Release(sheets); sheets = null;
            Release(book); book = null;
            //Release(books); books = null;
            Release(app); app = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public string? this[string cellName]
        {
            get => sheet?.Range[cellName].Value2.ToString();
            //{
            //    var range = sheet?.Range[cellName];
            //    return range?.Value2.ToString();
            //}
            set 
            {
                if(sheet != null) 
                    sheet.Range[cellName].Value2 = value;
                //range.Value = value;

            }
        }

        public void AddTable<T>(List<T> dataList)
        {
            if (dataList == null || dataList.Count == 0)
                return;

            var properties = typeof(T).GetProperties();
            for (int i = 0; i < properties.Length; i++)
            {
                sheet.Cells[1, i + 1] = properties[i].Name;
            }

            for (int i = 0; i < dataList.Count; i++)
            {
                for (int j = 0; j < properties.Length; j++)
                {
                    sheet.Cells[i + 2, j + 1] = properties[j].GetValue(dataList[i]);
                }
            }
        }

        public string? this[int row,int col]
        {
            get => sheet?.Cells[row,col].Value2.ToString();
            set
            {
                if (sheet != null)
                    sheet.Cells[row, col].Value = value;

            } //=> sheet.Cells[row,col].Value = value;
        }

        private void Release(object? obj)
        {
            if (obj != null)
                _ = Marshal.FinalReleaseComObject(obj);
        }
    }
}
