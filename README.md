
[Luban](https://luban.doc.code-philosophy.com/) \_\_tables\_\_.xlsx 一键更新工具

在Luban的使用中，所有用到的表都需要添加到__tables__.xlsx文件中，维护起来较为繁琐。虽然官方提供了[自动导入功能](https://luban.doc.code-philosophy.com/docs/beginner/importtable)，但仍然有一些限制，如表名无法使用中文、每个Excel文件只能对应一个表，一个文件包含多个工作表的情况无法使用。

使用本工具可将指定目录下的表格一键更新到__tables__.xlsx文件中，支持中文表名和多工作表。

### 使用方法

在Release页面下载最新版本，使用指令：

```bash
dotnet LubanHelper.dll updateTables --tablesPath *__tables__.xlsx路径* --dataPath *表文件目录*
```

例如：

**点我更新tables.bat**

```bash
set LUBAN_HELPER_DLL=.\LubanHelper\LubanHelper.dll

dotnet %LUBAN_HELPER_DLL% updateTables ^
    --tablesPath .\Data\__tables__.xlsx ^
    --dataPath .

pause
```

**点我更新tables.sh**

```bash
#!/bin/bash

LUBAN_HELPER_DLL=./LubanHelper/LubanHelper.dll

dotnet $LUBAN_HELPER_DLL updateTables \
    --tablesPath ./Data/__tables__.xlsx \
    --dataPath .

pause
```

表文件命名可使用中文，加`__`前缀忽略该文件，例如`__临时表.xlsx`会被忽略。

工作表(Sheet)需要以`模块名.类型名#表模式`的格式命名。模块名可省略；表模式对应__tables__.xlsx文件中的模式列，可用`#one`单例表、`#list`列表表、`#map`键值对表，表模式可省略，默认为map。同样加`__`前缀会忽略该工作表。

工作表命名示例：`LevelConfig`、`Shop.ItemConfig`、`GlobalConfig#one`、`Shop.RewardConfig#list`

结合Unity的使用示例可参考[Luban使用示例](https://github.com/PamisuMyon/pamisu-kit-unity/tree/main/samples/LubanExample)。
