
## Luban \_\_tables\_\_.xlsx 一键更新

在[Luban](https://luban.doc.code-philosophy.com/)的使用中，所有用到的表都需要添加到__tables__.xlsx文件中，维护起来较为繁琐。虽然官方提供了[自动导入功能](https://luban.doc.code-philosophy.com/docs/beginner/importtable)，但仍然有一些限制，如表名无法使用中文、每个Excel文件只能对应一个表，一个文件包含多个工作表的情况无法使用。

使用本工具可将指定目录下的表格一键更新到__tables__.xlsx文件中，支持中文表名和多工作表。

### 使用方法

在Release页面下载最新版本，使用指令：

```bash
dotnet LubanHelper.dll updateTables --tablesPath __tables__.xlsx路径 --dataPath 表文件目录
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

## 国际化/本地化一键填充

一种常见的动态国际化/本地化方式是建立一个表格，用本地化Id索引不同语言下的文本/资源，开发过程中，每有一个文本/资源就需要在表格中新增一行，这种手动维护的方式较为繁琐且容易出错。

使用本工具可以按规则检索所有配置表中的文本/资源备注列，将其填写到本地化表格中的默认语言列中，并生成其本地化ID，回填到原配置表的文本/资源Id列中。

本地化表格式：

![](https://raw.githubusercontent.com/PamisuMyon/gh-assets/main/images/pku/l10n_table.png)

其他配置表中需要有成对的备注列与Id列，如“名称备注”与“名称文本Id”、“插画资源备注”与“插画资源Id”：

![](https://raw.githubusercontent.com/PamisuMyon/gh-assets/main/images/pku/l10n_config_table.png)

备注列与Id列的命名需要有相同的名称与固定的后缀才能被工具识别， 例如“name_note”与“name_text_id”，名称为“name”，后缀“_note”表示其为备注列，“_text_id”表示其为Id列，后缀可以通过命令行参数自定义。

备注列可以加“#”前缀避免被Luban导出。

填写数据时只需要填写备注列，Id列留空，填写完毕后关闭所有表格，使用指令：

```bash
dotnet LubanHelper.dll updateL10N ^
    --l10nPath 本地化表路径 ^
    --dataPath 配置表目录 ^
    --noteColumnSuffix 备注列后缀 ^
    --textIdColumnSuffix Id列后缀  ^
    --l10nStartId 本地化起始Id
```

例如：

**点我更新本地化.bat**

```bash
chcp 65001
set LUBAN_HELPER_DLL=.\LubanHelper\LubanHelper.dll

dotnet %LUBAN_HELPER_DLL% updateL10N ^
    --l10nPath .\本地化.xlsx ^
    --dataPath . ^
    --noteColumnSuffix _note ^
    --textIdColumnSuffix _text_id ^
    --l10nStartId 20001

pause
```

指定本地化表文件为“本地化.xlsx”，配置表目录为当前文件夹，备注列后缀为“_note”，Id列后缀为“_text_id”，自动生成的本地化Id从20001开始。

运行结果：

![](https://raw.githubusercontent.com/PamisuMyon/gh-assets/main/images/pku/l10n_result.png)

后续有新增内容或修改，只需要再次执行即可，工具会自动识别，本地化表内没有的将会新增，已存在的会被复用。

具体使用示例可参考[Luban使用示例](https://github.com/PamisuMyon/pamisu-kit-unity/tree/main/samples/LubanExample)。
