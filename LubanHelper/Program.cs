using OfficeOpenXml;
using LubanHelper;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Please input arguments.");
            return;
        }

        if (args[0].Equals("syncTables"))
        {
            string tablesPath = null;
            string dataPath = null;
            
            for (int i = 1; i < args.Length; i++)
            {
                if (args[i] == "--tablesPath" && i + 1 < args.Length)
                {
                    tablesPath = args[i + 1];
                }
                if (args[i] == "--dataPath" && i + 1 < args.Length)
                {
                    dataPath = args[i + 1];
                }
            }

            if (string.IsNullOrEmpty(tablesPath))
            {
                Console.Error.WriteLine($"Error: tablesPath is null");
                return;
            }
            if (string.IsNullOrEmpty(dataPath))
            {
                Console.Error.WriteLine($"Error: dataPath is null");
                return;
            }
            
            SyncTables(tablesPath, dataPath);
                
        }
    }

    private static void SyncTables(string tablesFilePath, string dataPath)
    {
        if (!File.Exists(tablesFilePath))
        {
            Console.Error.WriteLine($"Error: Table file not exists: {tablesFilePath}");
            return;
        }
        if (!Directory.Exists(dataPath))
        {
            Console.Error.WriteLine($"Error: Data directory not exists: {tablesFilePath}");
            return;
        }

        var tableItems = new List<TableItem>();
        var files = Directory.GetFiles(dataPath);
        foreach (var file in files)
        {
            var info = new FileInfo(file);
            if (!info.Extension.Equals(".xlsx"))
                continue;
            using var package = new ExcelPackage(info);
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                Console.WriteLine($"Processing file {info.Name} sheet {worksheet.Name}");
                var tableItem = new TableItem
                {
                    FullName = $"Tb{worksheet.Name}",
                    ValueType = worksheet.Name,
                    Input = $"../{worksheet.Name}@{info.Name}",
                    Mode = worksheet.Name.Contains("Global")? "one" : "",
                };
                tableItems.Add(tableItem);
            }
        }

        var tablesFileInfo = new FileInfo(tablesFilePath);
        try
        {
            var backupFilePath = Path.Combine(tablesFileInfo.DirectoryName, $"{tablesFileInfo.Name.Replace(tablesFileInfo.Extension, "")}.backup{tablesFileInfo.Extension}");
            File.Copy(tablesFilePath, backupFilePath, true);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: Backup {tablesFileInfo.Name} failed: \n{ex.Message}");
        }

        try
        {
            using var tablesPackage = new ExcelPackage(tablesFileInfo);
            var worksheet = tablesPackage.Workbook.Worksheets[0];
            worksheet.DeleteRow(4, 200);
            var row = 4;
            foreach (var tableItem in tableItems)
            {
                worksheet.Cells[row, 2].Value = tableItem.FullName;
                worksheet.Cells[row, 3].Value = tableItem.ValueType;
                worksheet.Cells[row, 4].Value = "TRUE";
                worksheet.Cells[row, 5].Value = tableItem.Input;
                worksheet.Cells[row, 7].Value = tableItem.Mode;
                row++;
            }

            tablesPackage.Save();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: Write {tablesFileInfo.Name} failed: \n{ex.Message}");
        }

    }
}