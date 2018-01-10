using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadFromExcel
{
    class ReadExcelFile: IReadExcelFile
    {
        Excel.Application app;
        Excel.Workbook workBook;

        public Excel.Range GetExcelRange(string path, int sheetNr)
        {
            app = new Excel.Application();
            workBook = app.Workbooks.Open(path);
            Excel._Worksheet workSheet = (Excel.Worksheet)workBook.Sheets[sheetNr];
            Excel.Range excelRange = workSheet.UsedRange;

            return excelRange;
        }

        public void Close()
        {
            workBook.Close();
            app.Quit();
        }
    }
}
