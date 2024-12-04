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
        else if (args[0].Equals("updateL10N"))
        {
            UpdateL10NParams p = new UpdateL10NParams
            {
                NoteColumnSuffix = "Note",
                TextIdColumnSuffix = "TextId",
                L10NStartId = -1,
            };

            for (int i = 1; i < args.Length; i++)
            {
                if (args[i] == "--l10nPath" && i + 1 < args.Length)
                {
                    p.L10NFilePath = args[i + 1];
                }
                if (args[i] == "--dataPath" && i + 1 < args.Length)
                {
                    p.DataPath = args[i + 1];
                }
                if (args[i] == "--noteColumnSuffix" && i + 1 < args.Length)
                {
                    p.NoteColumnSuffix = args[i + 1];
                }
                if (args[i] == "--textIdColumnSuffix" && i + 1 < args.Length)
                {
                    p.TextIdColumnSuffix = args[i + 1];
                }
                if (args[i] == "--l10nStartId" && i + 1 < args.Length)
                {
                    if (int.TryParse(args[i + 1], out var id))
                        p.L10NStartId = id;
                }
                if (args[i] == "--appendFile" && i + 1 < args.Length)
                {
                    p.AppendFilePath = args[i + 1];
                }
            }
            
            UpdateL10N(p);
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
            Console.Error.WriteLine($"Error: Data directory not exists: {dataPath}");
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
        var uriA = new Uri(EnsureTrailingSlash(Path.GetFullPath(pathA)));
        var uriB = new Uri(EnsureTrailingSlash(Path.GetFullPath(pathB)));
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
    
    private static void UpdateL10N(UpdateL10NParams p)
    {
        if (!File.Exists(p.L10NFilePath))
        {
            Console.Error.WriteLine($"Error: L10N Table file not exists: {p.L10NFilePath}");
            return;
        }
        if (!Directory.Exists(p.DataPath))
        {
            Console.Error.WriteLine($"Error: Data directory not exists: {p.DataPath}");
            return;
        }

        // Fill L10N dict
        var l10nFileInfo = new FileInfo(p.L10NFilePath);
        using var l10nExcel = new ExcelPackage(l10nFileInfo);
        var l10nDict = new Dictionary<int, string>();
        var l10nDict1 = new Dictionary<string, int>();
        var l10nSheet = l10nExcel.Workbook.Worksheets[0];
        var l10nRow = GetContentStartRow(l10nSheet);
        var l10nId = 0;
        while (l10nRow < 10000) // hard-coded row limit
        {
            var idCell = l10nSheet.Cells[l10nRow, 2];   // hard-coded id column
            if (IsCellEmpty(idCell))
                break;
            
            try
            {
                l10nId = idCell.GetValue<int>();
                var content = l10nSheet.Cells[l10nRow, 3].Text; // hard-coded content column
                l10nDict.Add(l10nId, content);
                Console.WriteLine($"Read L10NTable row {l10nRow}: {l10nId} {content}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: Read {l10nFileInfo.Name} row {l10nRow} failed: \n{ex.Message}");
                return;
            }
            l10nRow++;
        }
        if (p.L10NStartId != -1 && l10nId < p.L10NStartId)
            l10nId = p.L10NStartId - 1;
        
        foreach (var pair in l10nDict)
        {
            if (l10nDict1.ContainsKey(pair.Value))
                continue;
            l10nDict1[pair.Value] = pair.Key;
        }
        
        // Process data files
        var files = Directory.GetFiles(p.DataPath);
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

            var excel = new ExcelPackage(info);
            foreach (var worksheet in excel.Workbook.Worksheets)
            {
                // Find L10N fields
                var noteIdDict = new Dictionary<ExcelRange, ExcelRange>();
                for (int i = 3; i < 1000; i++) // hard-coded column limit
                {
                    var cell = worksheet.Cells[1, i];
                    if (IsCellEmpty(cell))
                        break;
                    if (cell.Text.EndsWith(p.NoteColumnSuffix))
                        noteIdDict.Add(cell, null);
                }
                for (int i = 3; i < 1000; i++)  // hard-coded column limit
                {
                    var cell = worksheet.Cells[1, i];
                    if (IsCellEmpty(cell))
                        break;
                    if (cell.Text.EndsWith(p.TextIdColumnSuffix))
                    {
                        var field = cell.Text.Replace(p.TextIdColumnSuffix, "");
                        foreach (var key in noteIdDict.Keys)
                        {
                            if (key.Text.Replace(p.NoteColumnSuffix, "").Replace("#", "") == field)
                            {
                                noteIdDict[key] = cell;
                                break;
                            }
                        }
                    }
                }
                if (noteIdDict.Count == 0)
                {
                    // Console.WriteLine($"{info.Name} {worksheet.Name} L10N fields not found.");
                    continue;
                }
                Console.WriteLine($"================================");
                Console.WriteLine($"{info.Name} {worksheet.Name} L10N fields: ");
                foreach (var pair in noteIdDict)
                {
                    if (pair.Value != null)
                        Console.WriteLine($"{pair.Key.Text} : {pair.Value.Text}");
                }
                
                // Fill L10N fields
                foreach (var pair in noteIdDict)
                {
                    if (pair.Value == null)
                        continue;
                    var keyCol = pair.Key.Start.Column;
                    var valueCol = pair.Value.Start.Column;
                    for (var row = GetContentStartRow(worksheet); row < GetContentTotalRow(worksheet, row); row++)
                    {
                        var noteCell = worksheet.Cells[row, keyCol];
                        if (IsCellEmpty(noteCell))
                            break;
                        var idCell = worksheet.Cells[row, valueCol];
                        var noteText = noteCell.Text;
                        if (l10nDict1.TryGetValue(noteText, out int id))
                        {
                            idCell.Value = id;
                            Console.WriteLine($"Set row {row} {pair.Key.Text} = {noteText} {pair.Value.Text} = {id} (Exist)");
                        }
                        else
                        {
                            l10nId++;
                            l10nSheet.Cells[l10nRow, 2].Value = l10nId;
                            l10nSheet.Cells[l10nRow, 3].Value = noteText;
                            l10nDict.Add(l10nId, noteText);
                            l10nDict1.Add(noteText, l10nId);
                            l10nRow++;
                            idCell.Value = l10nId;
                            Console.WriteLine($"Set row {row} {pair.Key.Text} = {noteText} {pair.Value.Text} = {l10nId} (New)");
                        }
                    }
                }
                l10nExcel.Save();
            }
            excel.Save();
        }
        
        // Append file
        if (string.IsNullOrEmpty(p.AppendFilePath))
            return;
        if (!File.Exists(p.AppendFilePath))
        {
            Console.Error.WriteLine($"Error: Append Table file not exists: {p.AppendFilePath}");
            return;
        }
        var appendIds = new List<int>();
        var appendContents = new List<string>();
        foreach (var line in File.ReadLines(p.AppendFilePath))
        {
            var split = line.Split(",");
            if (!int.TryParse(split[0], out var id))
                continue;
            appendIds.Add(id);
            appendContents.Add(split[1]);
        }
        if (appendIds.Count == 0)
            return;
        Console.WriteLine($"================================");
        var appendRow = GetContentStartRow(l10nSheet);
        while (appendRow < 10000) // hard-coded row limit
        {
            var idCell = l10nSheet.Cells[appendRow, 2];   // hard-coded id column
            if (IsCellEmpty(idCell))
                break;
            try
            {
                l10nId = idCell.GetValue<int>();
                if (l10nId >= appendIds[0])
                    break;
            }
            catch (Exception ex)
            {
                return;
            }
            appendRow++;
        }
        for (int i = 0; i < appendIds.Count; i++)
        {
            Console.WriteLine($"{appendRow} {appendIds[i]} {appendContents[i]}");
            l10nSheet.Cells[appendRow, 2].Value = appendIds[i];
            l10nSheet.Cells[appendRow, 3].Value = appendContents[i];
            appendRow++;
        }
        l10nExcel.Save();
        Console.WriteLine($"{appendIds.Count} rows appended");
    }

    private static bool IsCellEmpty(ExcelRange cell)
    {
        return cell.Value == null || string.IsNullOrEmpty(cell.Text) || string.IsNullOrWhiteSpace(cell.Text);
    }

    private static int GetContentStartRow(ExcelWorksheet worksheet)
    {
        for (int i = 1; i < 50; i++)
        {
            var cell = worksheet.Cells[i, 1];
            if (IsCellEmpty(cell))
                return i;
            if (!cell.Text.StartsWith("#"))
                return i;
        }
        return 4;
    }

    private static int GetContentTotalRow(ExcelWorksheet worksheet, int contentStartRow)
    {
        for (int i = contentStartRow; i < 10000; i++)
        {
            var cell = worksheet.Cells[i, 2];
            if (IsCellEmpty(cell))
                return i;
        }
        return contentStartRow;
    }
}


public struct UpdateL10NParams
{
    public string L10NFilePath;
    public string DataPath;
    public string NoteColumnSuffix;
    public string TextIdColumnSuffix;
    public int L10NStartId;
    public bool ClearL10N; // not implemented yet
    public string AppendFilePath;
}
