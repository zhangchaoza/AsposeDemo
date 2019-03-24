using Aspose.Cells;
using System;
using System.IO;

namespace First
{
    class Program
    {
        static void Main(string[] args)
        {
            // For complete examples and data files, please go to https://github.com/aspose-cells/Aspose.Cells-for-.NET
            // The path to the documents directory.
            string dataDir = AppContext.BaseDirectory;

            // Save the Excel file.
            Gen().Save(dataDir + "MyBook_out.xlsx", SaveFormat.Xlsx);

            using (FileStream fs = new FileStream(dataDir + "MyBook_out2.xlsx", FileMode.Create))
            {
                Gen().Save(fs, SaveFormat.Xlsx);
            }

            using (MemoryStream ms = new MemoryStream(5000))
            {
                Gen().Save(ms, SaveFormat.Xlsx);
                byte[] buffer = new byte[ms.Length];
                ms.Seek(0, SeekOrigin.Begin);
                ms.Read(buffer, 0, buffer.Length);
                File.WriteAllBytes(dataDir + "MyBook_out3.xlsx", buffer);
            }
        }

        static Workbook Gen()
        {
            // Instantiate a Workbook object that represents Excel file.
            Workbook wb = new Workbook();

            // When you create a new workbook, a default "Sheet1" is added to the workbook.
            Worksheet sheet = wb.Worksheets[0];

            int currentRowIndex = 0;

            // Access the "A1" cell in the sheet.
            Cell cell = sheet.Cells["A1"];

            // Input the "Hello World!" text into the "A1" cell
            cell.PutValue("Hello World!");

            currentRowIndex++;

            Random r = new Random();
            for (int i = 0; i < 100; i++)
            {
                for (int j = 0; j < 7; j++)
                {
                    sheet.Cells[currentRowIndex + i, j].PutValue(r.NextDouble());
                }
            }

            return wb;
        }
    }
}
