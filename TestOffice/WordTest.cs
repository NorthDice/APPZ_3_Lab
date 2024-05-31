using OfficeLib;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestOffice
{
    [TestClass]
    public class WordTest
    {
        static int c1 = 0, c3 = 0;

        [TestMethod("01.Start / stop word com server")]
        public void Test01()
        {
            bool Result;

            int c2;

            c1 = Process.GetProcessesByName("WINWORD").Length;

            using (var x = new WordDocument()) ;
            {
                c2 = Process.GetProcessesByName("WINWORD").Length;
            }

            Result = (c2 - c1 == 1);

            Assert.IsTrue(Result);
        }

        [TestMethod("02.Create new word file")]
        public void Test02()
        {
            string FileName = "TestDocument";
            string FullName = $"{Directory.GetCurrentDirectory()}\\{FileName}.docx";

            if (File.Exists(FullName))
            {
                File.Delete(FullName);
            }

            using (var x = new WordDocument())
            {
                x[1] = "Paragraph 1";
                x[2] = "Paragraph 2";
                x[3] = "Paragraph 3";
                x[4] = "Paragraph 4";
                x[5] = "Paragraph 5";
                x[6] = "Paragraph 6";

                x.SaveAs(FullName);
            }

            Assert.IsTrue(File.Exists(FullName));
        }

        [TestMethod("03.Check content")]
        public void Test03()
        {
            string FileName = "TestDocument";
            string FullName = $"{Directory.GetCurrentDirectory()}\\{FileName}.docx";

            bool Result = File.Exists(FullName);

            using (var x = new WordDocument(FullName))
            {
                Result &= x[1] == "Paragraph 1";
                Result &= x[2] == "Paragraph 2";
                Result &= x[3] == "Paragraph 3";
                Result &= x[4] == "Paragraph 4";
                Result &= x[5] == "Paragraph 5";
                Result &= x[6] == "Paragraph 6";
            }

            Assert.IsTrue(Result);
        }

        [TestMethod("04.Check Garbage Collector")]
        public void Test04()
        {
            Thread.Sleep(3000);
            c3 = Process.GetProcessesByName("WINWORD").Length;

            Assert.IsTrue(c1 == c3);
        }
    }
}
