using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadFromExcel
{
    interface IReadExcelFile
    {
        Microsoft.Office.Interop.Excel.Range GetExcelRange(string path, int sheetNr);

        void Close();
    }
}
