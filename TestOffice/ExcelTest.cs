using OfficeLib;
using System.Diagnostics;

namespace TestOffice
{
    [TestClass]
    public class ExcelTest
    {
        static int c1 = 0, c3 = 0;
        [TestMethod("01.Start / stop excel com server")]
        public void Test01()
        {
            bool Result;

            int  c2;
            
            c1 = Process.GetProcessesByName("excel").Length;

            using (var x = new ExcelDocument());
            {
                c2 = Process.GetProcessesByName("excel").Length;
            }

            //c3 = Process.GetProcessesByName("excel").Length;

            Result =  (c2 - c1 == 1);

            Assert.IsTrue(Result);
        }
        [TestMethod("02.Create new excel file")]
        public void Test02()
        {
            string FileName = "TestDocument";
            string FullName = $"{Directory.GetCurrentDirectory()}\\{FileName}.xlsx";

            if (File.Exists(FullName))
            {
                File.Delete(FullName);
            }

            using (var x = new ExcelDocument())
            {
                x["A1"] = "Cell A1";
                x[2, 2] = "Cell 2,2";
                x[3, 3] = "100.0";
                x[4,4] = "25";
                x[5,5] = "27";

                x.SaveAs(FullName);
            }

            Assert.IsTrue(File.Exists(FullName));

        }
        [TestMethod("03.Check content")]
        public void Test03()
        {
            string FileName = "TestDocument";
            string FullName = $"{Directory.GetCurrentDirectory()}\\{FileName}.xlsx";

            bool Result = File.Exists(FullName);

            using (var x = new ExcelDocument(FullName))
            {
               Result &= x["A1"] == "Cell A1";
               Result &= x[2, 2] == "Cell 2,2";
               Result &= x[3, 3] == "100";
               Result &= x[4, 4] == "25";
               Result &= x[5, 5] =="27";
            }

            Assert.IsTrue(Result);

        }

        [TestMethod("04.Check Garbage Collector")]
        public void Test04()
        {
            Thread.Sleep(3000);
            c3 = Process.GetProcessesByName("excel").Length;

            Assert.IsTrue(c1 == c3);
        }
    }
}
