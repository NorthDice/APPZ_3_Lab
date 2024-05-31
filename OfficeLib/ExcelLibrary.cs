using System;

namespace OfficeLib
{
    public class ExcelLibrary
    {
        public static ExcelDocument CreateExcelDocument()
        {
            return new ExcelDocument();
        }

        public static ExcelDocument OpenExcelDocument(string FileName)
        {
            return new ExcelDocument(FileName);
        }
    }
}
