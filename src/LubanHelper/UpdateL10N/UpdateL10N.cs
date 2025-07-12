using OfficeOpenXml;

namespace LubanHelper.UpdateL10N;

// 本地化参数结构体
public struct UpdateL10NParams
{
    public string L10NFilePath; // 本地化表格路径
    public string DataPath; // 数据目录
    public string NoteColumnSuffix; // 注释列后缀
    public string TextIdColumnSuffix; // 文本ID列后缀
    public int L10NStartId; // 本地化起始ID
    public bool ClearL10N; // 是否清空本地化，未实现
    public string AppendFilePath; // 追加文件路径
}

public static class UpdateL10N
{
    
    public static void UpdateL10NHandle(string[] args)
    {
        // 初始化本地化参数
        UpdateL10NParams p = new UpdateL10NParams
        {
            NoteColumnSuffix = "Note", // 注释列后缀
            TextIdColumnSuffix = "TextId", // 文本ID列后缀
            L10NStartId = -1, // 本地化起始ID
        };

        // 解析参数
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
            
        Update(p);
    }
    
    // 更新本地化表格
    private static void Update(UpdateL10NParams p)
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

        // 填充L10N字典
        var l10nFileInfo = new FileInfo(p.L10NFilePath);
        using var l10nExcel = new ExcelPackage(l10nFileInfo);
        var l10nDict = new Dictionary<int, string>(); // id->文本
        var l10nDict1 = new Dictionary<string, int>(); // 文本->id
        var l10nSheet = l10nExcel.Workbook.Worksheets[0];
        var l10nRow = l10nSheet.GetContentStartRow();
        var l10nId = 0;
        while (l10nRow < Utils.MaxRowCount) // 行数上限
        {
            var idCell = l10nSheet.Cells[l10nRow, 2];   // id列
            if (idCell.IsCellEmpty())
                break;
            try
            {
                l10nId = idCell.GetValue<int>();
                var content = l10nSheet.Cells[l10nRow, 3].Text; // 文本内容列
                l10nDict.Add(l10nId, content);
                Console.WriteLine($"Read L10NTable row {l10nRow}: {l10nId} {content}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: Read {l10nFileInfo.Name} row {l10nRow} failed: \n{ex}");
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
        
        // 处理数据文件
        var files = Directory.GetFiles(p.DataPath);
        foreach (var file in files)
        {
            var info = new FileInfo(file);
            if (!info.Extension.Equals(".xlsx"))
                continue;
            if (Utils.IsSkipFile(info.Name))
            {
                Console.WriteLine($"Skip file {info.Name}");
                continue;
            }

            using var excel = new ExcelPackage(info);
            foreach (var worksheet in excel.Workbook.Worksheets)
            {
                // 查找本地化字段
                var noteIdDict = new Dictionary<ExcelRange, ExcelRange>();
                for (int i = 3; i < Utils.MaxColCount; i++) // 列数上限
                {
                    var cell = worksheet.Cells[1, i];
                    if (cell.IsCellEmpty())
                        break;
                    if (cell.Text.EndsWith(p.NoteColumnSuffix))
                        noteIdDict.Add(cell, null);
                }
                for (int i = 3; i < Utils.MaxColCount; i++)  // 列数上限
                {
                    var cell = worksheet.Cells[1, i];
                    if (cell.IsCellEmpty())
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
                
                // 填充本地化字段
                foreach (var pair in noteIdDict)
                {
                    if (pair.Value == null)
                        continue;
                    var keyCol = pair.Key.Start.Column;
                    var valueCol = pair.Value.Start.Column;
                    for (var row = worksheet.GetContentStartRow(); row < worksheet.GetContentTotalRow(row); row++)
                    {
                        var noteCell = worksheet.Cells[row, keyCol];
                        if (noteCell.IsCellEmpty())
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
        
        // 追加文件内容
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
        var appendRow = l10nSheet.GetContentStartRow();
        while (appendRow < Utils.MaxRowCount) // 行数上限
        {
            var idCell = l10nSheet.Cells[appendRow, 2];   // id列
            if (idCell.IsCellEmpty())
                break;
            try
            {
                l10nId = idCell.GetValue<int>();
                if (l10nId >= appendIds[0])
                    break;
            }
            catch (Exception)
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
}