using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeLib
{
    public class WordDocument : IDisposable
    {
        Word.Application? app = null;
        Word.Document? doc = null;

        public WordDocument()
        {
            app = new Word.Application();
            doc = app.Documents.Add();
            app.Visible = false;
        }

        public WordDocument(string FileName)
        {
            app = new Word.Application();
            doc = app.Documents.Open(FileName);
            app.Visible = false;
        }

        public void SaveAs(string FileName)
        {
            doc?.SaveAs2(FileName);
        }

        public void Dispose()
        {
            doc?.Close(false);
            app?.Quit();

            Release(doc); doc = null;
            Release(app); app = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public string? this[int index]
        {
            get
            {
                if (doc != null && index > 0 && index <= doc.Paragraphs.Count)
                {
                    return doc.Paragraphs[index].Range.Text.Trim();
                }
                return null;
            }
            set
            {
                if (doc != null && value != null)
                {
                    if (index > 0 && index <= doc.Paragraphs.Count)
                    {
                        doc.Paragraphs[index].Range.Text = value;
                    }
                    else
                    {
                        while (doc.Paragraphs.Count < index)
                        {
                            doc.Paragraphs.Add();
                        }
                        doc.Paragraphs[index].Range.Text = value;
                    }
                }
            }
        }

        public void CreateTable(int numRows, int numCols)
        {
            if (numRows <= 0 || numCols <= 0)
            {
                throw new ArgumentException("Number of rows and columns must be greater than zero.");
            }

            // Вставляем пустой абзац в конце документа
            Word.Paragraph para = doc.Content.Paragraphs.Add();
            Word.Range range = para.Range;

            // Создаем таблицу в новом пустом абзаце
            Word.Table table = doc.Tables.Add(range, numRows, numCols);

            // Настраиваем границы таблицы
            table.Borders.Enable = 1;
            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth050pt;
            table.Borders.OutsideColor = Word.WdColor.wdColorBlack;

            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineWidth = Word.WdLineWidth.wdLineWidth050pt;
            table.Borders.InsideColor = Word.WdColor.wdColorBlack;
        }



        public void AddDataToCell(int row, int col, string data)
        {
            if (doc.Tables.Count > 0 && row <= doc.Tables[1].Rows.Count && col <= doc.Tables[1].Columns.Count)
            {
                doc.Tables[1].Cell(row, col).Range.Text = data;
            }
        }

        public void AddTable<T>(List<T> dataList)
        {
            if (dataList == null || dataList.Count == 0)
                return;

            var properties = typeof(T).GetProperties();
            CreateTable(dataList.Count + 1, properties.Length);

            for (int i = 0; i < properties.Length; i++)
            {
                AddDataToCell(1, i + 1, properties[i].Name);
            }

            for (int i = 0; i < dataList.Count; i++)
            {
                for (int j = 0; j < properties.Length; j++)
                {
                    AddDataToCell(i + 2, j + 1, properties[j].GetValue(dataList[i])?.ToString());
                }
            }
        }

        private void Release(object? obj)
        {
            if (obj != null)
                _ = Marshal.FinalReleaseComObject(obj);
        }
    }
}
