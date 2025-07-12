using System.Text.RegularExpressions;
using OfficeOpenXml;
namespace LubanHelper;

public static class Utils
{
    // 魔法数字常量
    public const int MaxRowCount = 10000;
    public const int MaxColCount = 1000;
    public const int MaxHeaderScanRow = 50;
    public const int DefaultContentStartRow = 4;

    // 排除文件名正则（以__开头的文件）
    private static Regex _excludeRegex = new("^__");
    
    public static bool IsSkipFile(string fileName)
    {
        return _excludeRegex.IsMatch(fileName);
    }
    
    // 获取相对路径
    public static string GetRelativePath(string pathA, string pathB)
    {
        var uriA = new Uri(EnsureTrailingSlash(Path.GetFullPath(pathA)));
        var uriB = new Uri(EnsureTrailingSlash(Path.GetFullPath(pathB)));
        var relativeUri = uriA.MakeRelativeUri(uriB);
        return Uri.UnescapeDataString(relativeUri.ToString());
    }
    
    // 确保路径以分隔符结尾
    public static string EnsureTrailingSlash(string path)
    {
        if (!path.EndsWith(Path.DirectorySeparatorChar.ToString()) && Directory.Exists(path))
        {
            return $"{path}{Path.DirectorySeparatorChar}";
        }
        return path;
    }
    
    // 判断单元格是否为空
    public static bool IsCellEmpty(ExcelRange cell)
    {
        return cell.Value == null || string.IsNullOrEmpty(cell.Text) || string.IsNullOrWhiteSpace(cell.Text);
    }
    
    // 获取内容起始行
    public static int GetContentStartRow(ExcelWorksheet worksheet)
    {
        for (int i = 1; i < MaxHeaderScanRow; i++)
        {
            var cell = worksheet.Cells[i, 1];
            if (IsCellEmpty(cell))
                return i;
            if (!cell.Text.StartsWith("#"))
                return i;
        }
        return DefaultContentStartRow;
    }
    
    // 获取内容总行数
    public static int GetContentTotalRow(ExcelWorksheet worksheet, int contentStartRow)
    {
        for (int i = contentStartRow; i < MaxRowCount; i++)
        {
            var cell = worksheet.Cells[i, 2];
            if (IsCellEmpty(cell))
                return i;
        }
        return contentStartRow;
    }

}