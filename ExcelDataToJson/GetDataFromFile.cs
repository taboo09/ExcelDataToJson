using ExcelDataToJson;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadFromExcel
{
    class GetDataFromFile
    {
        private IReadExcelFile excelFile;
        private Excel.Range excelRange;
        private List<Inventory> Inventorylist;

        public GetDataFromFile(string path, int sheetNr)
        {
            excelFile = new ReadExcelFile();
            excelRange = excelFile.GetExcelRange(path, sheetNr);
            Inventorylist = new List<Inventory>();
        }

        public List<Inventory> GetObjectFromExcel()
        {
            int colsCount = excelRange.Columns.Count;
            int rowsCount = excelRange.Rows.Count;

            // Save entire Excel row to a list of strings
            var listString = new List<string>();

            // Excel indexing starts from 1
            for (int i = 1; i <= rowsCount; i++)
            {
                for (int j = 1; j <= colsCount; j++)
                {
                    if (excelRange.Cells[i, j].Value2 == null)
                    {
                        listString.Add("Undefined");
                    }
                    else listString.Add((excelRange.Cells[i, j].Value2).ToString());
                }
                var Inventory = new Inventory();

                Inventory.Brand = listString[0];
                Inventory.ProductCode = listString[1];
                Inventory.Description = listString[2];
                Inventory.Stock = Convert.ToInt32(listString[3]);

                // Add object to the list of objects
                Inventorylist.Add(Inventory);

                listString.Clear();

            }

            // Close file
            excelFile.Close();

            return Inventorylist;
        }
    }
}
