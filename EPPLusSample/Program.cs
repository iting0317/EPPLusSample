using EPPLusSample;

XLSHelper xlsHeler = new XLSHelper();

string path = @"..\..\..\測試匯入資料.xlsx";
string fullPath = Path.Combine(Directory.GetCurrentDirectory(), path);

var valid = xlsHeler.ValidFile(fullPath);
if (!valid.IsValid)
{
    Console.WriteLine(valid.Msg);
    Console.ReadKey();
    Environment.Exit(0);
}

var fileStream = File.OpenRead(fullPath);
            
//讀取為List<string>
var stringResult = xlsHeler.ReadExcelToStringList(fileStream);
stringResult.ForEach(line => Console.WriteLine(line));

//讀取為DataTable
var datatableResult = xlsHeler.ReadExcelToDataTable(fileStream);
