using System;
using System.Collections.Generic;
using MainConsole;
using OfficeLib;

class Program
{
    static void Main()
    {
        
        var sessions = new List<Session>
        {
            new Session { Id = 1, Name = "Session 1", StartTime = DateTime.Now, EndTime = DateTime.Now.AddHours(1) },
            new Session { Id = 2, Name = "Session 2", StartTime = DateTime.Now.AddHours(1), EndTime = DateTime.Now.AddHours(2) }
        };

      
        var users = new List<User>
        {
            new User { Id = 1, FirstName = "John", LastName = "Doe", Email = "john.doe@example.com" },
            new User { Id = 2, FirstName = "Jane", LastName = "Smith", Email = "jane.smith@example.com" }
        };


        using (var excelDoc = ExcelLibrary.CreateExcelDocument())
        {
            excelDoc.AddTable(sessions);
            excelDoc.SaveAs("session.xlsx");
            excelDoc.AddTable(users);
            excelDoc.SaveAs("users.xlsx");
        }

        
        using (var wordDoc = WordLibrary.CreateWordDocument())
        {
            wordDoc.AddTable(sessions);
            wordDoc.SaveAs("session.docx");
            wordDoc.AddTable(users);
            wordDoc.SaveAs("users.docx");
        }

        Console.WriteLine("Success.");
    }
}
