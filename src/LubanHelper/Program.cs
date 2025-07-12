using System.Text.RegularExpressions;
using OfficeOpenXml;
using LubanHelper;
using LubanHelper.UpdateTables;

class Program
{
    static void Main(string[] args)
    {
        // 主入口，解析命令行参数
        if (args.Length == 0)
        {
            Console.WriteLine("Please input arguments.");
            return;
        }

        if (args[0].Equals("updateTables"))
        {
            UpdateTables.UpdateTablesHandle(args);
        }
        else if (args[0].Equals("updateL10N"))
        {
            
        }
    }

}