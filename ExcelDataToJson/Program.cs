using ReadFromExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            var getDataFromFile = new GetDataFromFile(@"Your Excel Path.xlsx", 1);



            List<Inventory> listOfInventory = getDataFromFile.GetObjectFromExcel();

            var listOfObjectToJson = new ListObjectToJSON(@"Json or txt path file.json");

            listOfObjectToJson.ListToJSON(listOfInventory);
        }
    }
}
