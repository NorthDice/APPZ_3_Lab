using System;

namespace OfficeLib
{
    public class WordLibrary
    {
        public static WordDocument CreateWordDocument()
        {
            return new WordDocument();
        }

        public static WordDocument OpenWordDocument(string FileName)
        {
            return new WordDocument(FileName);
        }
    }
}
