using System.Data;
using System.Text.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace AlarmPeople;
class Program
{
    static void Main(string[] args)
    {
        var inp = args.Length > 1 ? args[0] : "Book1.xlsx";
        var outp = args.Length > 2 ? args[1] : "Book1.json";

        using (var outs = new FileStream(outp, FileMode.Create))
        {
            ReadExcel(inp, outs);
        }

        Console.WriteLine("DONE...");
    }

    static void ReadExcel(string input_filename, Stream output_stream)
    {
        DataTable dtTable = new DataTable();
        List<string> rowList = new List<string>();
        ISheet sheet;
        HashSet<int> ignored_columns = new HashSet<int>();
        using (var stream = new FileStream(input_filename, FileMode.Open))
        {
            stream.Position = 0;
            XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
            sheet = xssWorkbook.GetSheetAt(0);
            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;
            for (int j = 0; j < cellCount; j++)
            {
                ICell cell = headerRow.GetCell(j);
                if (cell == null || string.IsNullOrWhiteSpace(cell.ToString()))
                {
                    ignored_columns.Add(j);
                }
                else
                {
                    dtTable.Columns.Add(cell.ToString());
                }
            }
            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) 
                {
                    continue;
                }
                if (row.Cells.All(d => d.CellType == CellType.Blank)) 
                {
                    continue;
                }
                for (int j = 0; j < cellCount; j++)
                {
                    if (ignored_columns.Contains(j))
                    {
                        continue; // dont get values for ignored columns.
                    }
                    var c = row.GetCell(j)?.ToString();
                    rowList.Add(string.IsNullOrEmpty(c) ? "NULL" : c.Trim());
                }
                if (rowList.Count > 0)
                {                    
                    dtTable.Rows.Add(rowList.ToArray());
                }                
                rowList.Clear();
            }
        }
        var res = DataTable_to_Json(dtTable);
        var buf = System.Text.Encoding.UTF8.GetBytes(res);
        if (buf != null) 
        {
            output_stream.Write(buf, 0, buf.Length);
        }       
    }

    public static string DataTable_to_Json(DataTable dataTable)
    {
        if (dataTable == null)
        {
            return string.Empty;
        }

        var data = dataTable.Rows.OfType<DataRow>()
                    .Select(row => dataTable.Columns.OfType<DataColumn>()
                        .ToDictionary(col => col.ColumnName, c => row[c]));

        return JsonSerializer.Serialize(data);
    }

}