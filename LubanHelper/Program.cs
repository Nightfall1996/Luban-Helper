using System.Text.RegularExpressions;
using OfficeOpenXml;
using LubanHelper;

class Program
{
    
    private static Regex _excludeRegex = new("^__");
    private static Regex _oneMode = new("#one$", RegexOptions.IgnoreCase);
    private static Regex _listMode = new("#list$", RegexOptions.IgnoreCase);
    private static Regex _mapMode = new("#map$", RegexOptions.IgnoreCase);
    private static Regex _part = new("-\\d+$");
    
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Please input arguments.");
            return;
        }

        if (args[0].Equals("updateTables"))
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
            
            UpdateTables(tablesPath, dataPath);
                
        }
    }

    private static void UpdateTables(string tablesFilePath, string dataPath)
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

        var relativePath = GetRelativePath(tablesFilePath, dataPath);
        var tableItemDict = new Dictionary<string, TableItem>();
        var files = Directory.GetFiles(dataPath);
        foreach (var file in files)
        {
            var info = new FileInfo(file);
            if (!info.Extension.Equals(".xlsx"))
                continue;
            if (_excludeRegex.IsMatch(info.Name))
            {
                Console.WriteLine($"Skip file {info.Name}");
                continue;
            }
            
            using var package = new ExcelPackage(info);
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                if (_excludeRegex.IsMatch(worksheet.Name))
                {
                    Console.WriteLine($"Skip sheet {worksheet.Name} of file {info.Name} ");
                    continue;
                }
                
                Console.WriteLine($"Processing sheet {worksheet.Name} of file {info.Name}");
                var sheetName = worksheet.Name;
                string mode = "";
                if (_oneMode.IsMatch(sheetName))
                {
                    sheetName = _oneMode.Replace(sheetName, "");
                    mode = "one";
                }
                else if (_listMode.IsMatch(sheetName))
                {
                    sheetName = _listMode.Replace(sheetName, "");
                    mode = "list";
                }
                else if (_mapMode.IsMatch(sheetName))
                {
                    sheetName = _mapMode.Replace(sheetName, "");
                }
                
                var split = sheetName.Split(".");
                var valueType = split[^1];
                if (_part.IsMatch(valueType))
                {
                    valueType = _part.Replace(valueType, "");
                }
                
                string fullName;
                if (split.Length > 1)
                    fullName = $"{string.Join(".", split.Take(split.Length - 1))}.Tb{valueType}";
                else
                    fullName = $"Tb{valueType}";

                if (tableItemDict.TryGetValue(fullName, out var itemInDict))
                {
                    // 多工作表对单数据表，追加到输入文件字段
                    itemInDict.Input = $"{itemInDict.Input},{relativePath}{worksheet.Name}@{info.Name}";
                }
                else
                {
                    var tableItem = new TableItem
                    {
                        FullName = fullName,
                        ValueType = valueType,
                        Input = $"{relativePath}{worksheet.Name}@{info.Name}",
                        Mode = mode
                    };
                    tableItemDict.Add(tableItem.FullName, tableItem);
                }
                
                // Console.WriteLine($"{tableItem.FullName} {tableItem.ValueType} {tableItem.Input} {tableItem.Mode}");
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
            worksheet.DeleteRow(4, 999);
            var row = 4;
            foreach (var tableItem in tableItemDict.Values)
            {
                worksheet.Cells[row, 2].Value = tableItem.FullName;
                worksheet.Cells[row, 3].Value = tableItem.ValueType;
                worksheet.Cells[row, 4].Value = "TRUE";
                worksheet.Cells[row, 5].Value = tableItem.Input;
                worksheet.Cells[row, 7].Value = tableItem.Mode;
                row++;
            }

            tablesPackage.Save();
            
            Console.WriteLine("Completed.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: Write {tablesFileInfo.Name} failed: \n{ex.Message}");
        }

    }

    private static string GetRelativePath(string pathA, string pathB)
    {
        var uriA = new Uri(EnsureTrailingSlash(pathA));
        var uriB = new Uri(EnsureTrailingSlash(pathB));
        var relativeUri = uriA.MakeRelativeUri(uriB);
        return Uri.UnescapeDataString(relativeUri.ToString());
    }
    
    private static string EnsureTrailingSlash(string path)
    {
        if (!path.EndsWith(Path.DirectorySeparatorChar.ToString()) && Directory.Exists(path))
        {
            return $"{path}{Path.DirectorySeparatorChar}";
        }
        return path;
    }

}