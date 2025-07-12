using System.Diagnostics;
using System.Text.RegularExpressions;
using OfficeOpenXml;
namespace LubanHelper.UpdateTables;

public class TableItem
{
    public string FullName;
    public string ValueType;
    public string Input;
    public string Mode;
}

public static class UpdateTables
{
    // 匹配#one结尾的表名，单行模式
    private static Regex _oneMode = new("#one$", RegexOptions.IgnoreCase);
    // 匹配#list结尾的表名，列表模式
    private static Regex _listMode = new("#list$", RegexOptions.IgnoreCase);
    // 匹配#map结尾的表名，映射模式
    private static Regex _mapMode = new("#map$", RegexOptions.IgnoreCase);
    // 匹配-数字结尾的表名片段
    private static Regex _part = new("-\\d+$");
    
    public static void UpdateTablesHandle(string[] args)
    {
        string tablesPath = null;
        string dataPath = null;
            
        // 解析参数
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
        
        Update(tablesPath, dataPath);
    }
    
    /// <summary>
    /// 更新表格配置，将指定数据目录下的所有 Excel 文件及其工作表信息，
    /// 解析为 TableItem 并写入到总表格配置文件（tablesFilePath）中。
    /// 
    /// 主要流程：
    /// 1. 校验总表格文件和数据目录是否存在。
    /// 2. 遍历数据目录下所有 .xlsx 文件，跳过无关文件。
    /// 3. 对每个工作表，解析表名、模式（one/list/map）、类型等，构建 TableItem 字典。
    /// 4. 支持多工作表合并到同一数据表（Input 字段追加）。
    /// 5. 备份原有总表格文件。
    /// 6. 清空总表格旧数据，从第4行起写入新的 TableItem 信息。
    /// 7. 捕获并输出备份、写入过程中的异常。
    /// </summary>
    /// <param name="tablesFilePath">总表格配置文件路径</param>
    /// <param name="dataPath">数据表目录路径</param>
    private static void Update(string tablesFilePath, string dataPath)
    {
        if (!File.Exists(tablesFilePath))
        {
            Console.Error.WriteLine($"Error: Table file not exists: {tablesFilePath}");
            return;
        }
        if (!Directory.Exists(dataPath))
        {
            Console.Error.WriteLine($"Error: Data directory not exists: {dataPath}");
            return;
        }

        var tableItemDict = new Dictionary<string, TableItem>();
        var files = Directory.GetFiles(dataPath, "*.xlsx", SearchOption.AllDirectories)
            .Where(file => 
                !Path.GetFileName(file).StartsWith("__") && 
                !Path.GetFileName(file).StartsWith("~$"))
            .ToArray();
        
        foreach (var file in files)
        {
            var info = new FileInfo(file);
            if (!info.Extension.Equals(".xlsx"))
                continue;
            
            if (Utils.IsSkipFile(info.Name))
            {
                // Console.WriteLine($"Skip file {info.Name}");
                continue;
            }
            
            // 遍历数据目录下所有 Excel 文件
            using var package = new ExcelPackage(info);
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                // 跳过无关工作表（如临时表、隐藏表等）
                if (Utils.IsSkipFile(worksheet.Name))
                {
                    // Console.WriteLine($"Skip sheet {worksheet.Name} of file {info.Name} ");
                    continue;
                }
                
                var relativeFilePath = Utils.GetRelativeFilePath(dataPath, file);
                // Console.WriteLine($"Processing sheet {worksheet.Name} of file {info.Name}. relativePath : {relativeFilePath}");
                // 解析工作表名，判断模式（one/list/map）
                var sheetName = worksheet.Name;
                string mode = "";
                if (_oneMode.IsMatch(sheetName))
                {
                    sheetName = _oneMode.Replace(sheetName, "");
                    mode = "one"; // 单行模式
                }
                else if (_listMode.IsMatch(sheetName))
                {
                    sheetName = _listMode.Replace(sheetName, "");
                    mode = "list"; // 列表模式
                }
                else if (_mapMode.IsMatch(sheetName))
                {
                    sheetName = _mapMode.Replace(sheetName, "");
                    // map模式
                }
                
                // 解析类型名，去除片段后缀
                var split = sheetName.Split(".");
                var valueType = split[^1];
                if (_part.IsMatch(valueType))
                {
                    valueType = _part.Replace(valueType, "");
                }
                
                // 组装表全名（如 TbXXX）
                string fullName = $"Tb{valueType}";
                if (tableItemDict.TryGetValue(fullName, out var itemInDict))
                {
                    // 多工作表对单数据表，追加到输入文件字段
                    itemInDict.Input = $"{itemInDict.Input},{worksheet.Name}@{relativeFilePath}";
                }
                else
                {
                    var tableItem = new TableItem
                    {
                        FullName = fullName,
                        ValueType = $"{valueType}Bean",
                        Input = $"{worksheet.Name}@{relativeFilePath}",
                        Mode = mode
                    };
                    tableItemDict.Add(tableItem.FullName, tableItem);
                }
                
                // Console.WriteLine($"{tableItem.FullName} {tableItem.ValueType} {tableItem.Input} {tableItem.Mode}");
            }
        }

        // 备份原表格
        var tablesFileInfo = new FileInfo(tablesFilePath);
        // try
        // {
        //     var backupFilePath = Path.Combine(tablesFileInfo.DirectoryName, $"{tablesFileInfo.Name.Replace(tablesFileInfo.Extension, "")}.backup{tablesFileInfo.Extension}");
        //     File.Copy(tablesFilePath, backupFilePath, true);
        // }
        // catch (Exception ex)
        // {
        //     Console.Error.WriteLine($"Error: Backup {tablesFileInfo.Name} failed: \n{ex.Message}");
        // }

        try
        {
            using var tablesPackage = new ExcelPackage(tablesFileInfo);
            var worksheet = tablesPackage.Workbook.Worksheets[0];
            worksheet.DeleteRow(4, Utils.MaxRowCount); // 清除旧数据
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
}