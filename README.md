# ExcelDataToJson 
### Serialize Excel Data to JSON Object in C#


Add the reference to Microsoft Excel XX.X Object Library, located in the COM tab of the Reference Manager to manipulate Microsoft Excel files:

```
using Excel = Microsoft.Office.Interop.Excel;
```


Create COM Objects. Create a COM object for everything that is referenced

```
app = new Excel.Application();
workBook = app.Workbooks.Open(path);
Excel._Worksheet workSheet = (Excel.Worksheet)workBook.Sheets[sheetNr];
Excel.Range excelRange = workSheet.UsedRange;
```

Note: Excel indexing starts from 1!

In my example I parsed the data into an object:

```
class Inventory
    {
        public string Brand { get; set; }
        public string ProductCode { get; set; }
        public string Description { get; set; }
        public int Stock { get; set; }
    }
 ```

Then serialized it into a Json object using Newtonsoft.Json Framework for .NET:

```
string JSON_Format = JsonConvert.SerializeObject(listOfObjectT);
```

Testing and working properly in .NET Framework version=v4.5.2
