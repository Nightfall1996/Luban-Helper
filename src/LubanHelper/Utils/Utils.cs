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
    
    public static string GetRelativeFilePath(string basePath, string fullFilePath)
    {
        // 规范化路径格式
        basePath = Path.GetFullPath(basePath);
        fullFilePath = Path.GetFullPath(fullFilePath);
        
        // 检查文件是否在基础目录下
        if (!fullFilePath.StartsWith(basePath, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("文件不在指定基础目录下");
        }
        
        // 计算相对路径
        string relativePath = fullFilePath.Substring(basePath.Length);
        
        // 移除开头的路径分隔符
        relativePath = relativePath.TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        
        return relativePath;
    }

    // 判断单元格是否为空
    public static bool IsCellEmpty(this ExcelRange cell)
    {
        return cell.Value == null || string.IsNullOrEmpty(cell.Text) || string.IsNullOrWhiteSpace(cell.Text);
    }
    
    // 获取内容起始行
    public static int GetContentStartRow(this ExcelWorksheet worksheet)
    {
        for (int i = 1; i < MaxHeaderScanRow; i++)
        {
            var cell = worksheet.Cells[i, 1];
            if (cell.IsCellEmpty())
                return i;
            if (!cell.Text.StartsWith("#"))
                return i;
        }
        return DefaultContentStartRow;
    }
    
    // 获取内容总行数
    public static int GetContentTotalRow(this ExcelWorksheet worksheet, int contentStartRow)
    {
        for (int i = contentStartRow; i < MaxRowCount; i++)
        {
            var cell = worksheet.Cells[i, 2];
            if (cell.IsCellEmpty())
                return i;
        }
        return contentStartRow;
    }

}