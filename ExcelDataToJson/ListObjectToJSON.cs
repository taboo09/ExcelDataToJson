using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadFromExcel
{
    public class ListObjectToJSON
    {
        private string filePath;

        public ListObjectToJSON(string _filePath)
        {
            filePath = _filePath;
            File.Delete(filePath);
        }

        public void ListToJSON<T>(List<T> listOfObjectT)
        {
            string JSON_Format = JsonConvert.SerializeObject(listOfObjectT);

            StreamWriter sw = new StreamWriter(filePath);
            sw.Write(JSON_Format);

            sw.Close();

            Console.WriteLine("List of Inventory succesfuly serialized to Json format!");
        }
    }
}
