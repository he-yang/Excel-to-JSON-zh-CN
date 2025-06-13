# 1. 快速开始
 <a name="Introduction"></a> 
[English](https://excel-to-json.wtsolutions.cn/latest/getstarted.html)

WTSolutions 的 Excel 转 JSON 是一款**Microsoft Excel 加载项**或**Web 应用程序**，可以将 Excel 转换为 JSON，支持转换扁平 JSON 和嵌套 JSON。


# 2. 系统要求
 <a name="Requirements"></a>
`选项 1. 在 Web 浏览器中加载 Excel 转 JSON`
* 支持 JavaScript 的 Web 浏览器，如 Google Chrome、Mozilla Firefox、Safari 或 Microsoft Edge。

`选项 2. 在 Excel 中旁加载 Excel 转 JSON`（推荐）
* Excel 2013 Service Pack 1 或更高版本，
* Excel 2016 for Mac，
* Excel 2016 或更高版本，
* Excel Online，
* Office 365 等。


# 3. 快速入门
<a name="Quickstarted"></a> 
本快速入门适用于 v3.0.0 版本


## 3.1 （旁）加载 Excel 转 JSON

`选项 1. 在 Web 浏览器中加载 Excel 转 JSON`
* 打开支持 JavaScript 的 Web 浏览器，如 Google Chrome、Mozilla Firefox、Safari 或 Microsoft Edge。
* 在 Web 浏览器中打开以下 URL：<a href="https://s.wtsolutions.cn/excel-to-json.html" target="_blank">https://s.wtsolutions.cn/excel-to-json.html</a>


`选项 2. 在 Excel 中旁加载 Excel 转 JSON`（推荐）
* 在 Excel 2013/2016、Excel Online 或 Office 365 中打开一个新的数据表。
* **主页**选项卡或**插入**选项卡 > 加载项
* 在搜索框中，输入“Excel to JSON”
* 按照屏幕上的说明安装加载项，您将在**主页**选项卡上看到一个“JSON 转 Excel”按钮。
* **主页**选项卡 > Excel 转 JSON > 转换
* 现在您已准备好使用此加载项。


 <a name="Useadd-in"></a> 

## 3.2 使用 Excel 转 JSON

* 准备您的 Excel 表格
* 通过以下两种方式之一加载您的 Excel 数据：
    1. `在 Web 浏览器中加载 Excel 转 JSON`：将 Excel 数据复制并粘贴到文本区域，或者
    2. `在 Excel 中旁加载 Excel 转 JSON`：直接从 Excel 工作表中选择您的数据。
* 设置转换设置
* 点击“开始”按钮
* 之后您将在“开始”按钮下方看到转换后的 JSON

### 准备您的 Excel 表格

有两种方式可以将 Excel 表格加载到 Excel 转 JSON 中：

* `在 Web 浏览器中加载 Excel 转 JSON`
    *  将 Excel 数据复制并粘贴到文本区域
    *  您可以复制并粘贴来自 Excel、Google Sheets 或任何其他 Excel 兼容软件的数据，数据之间用 Tab 分隔
    *  您也可以复制并粘贴逗号分隔的 CSV 数据
* `在 Excel 中旁加载 Excel 转 JSON`：直接从 Excel 工作表中选择您的数据。

### 输出 JSON 导出

有几种方法可以将生成的 JSON 保存到您的本地计算机。

* `复制和粘贴`。JSON 生成后，您将在文本区域中看到它们，您可以简单地将它们复制并粘贴到任何您想要的地方。
* `复制到剪贴板`。JSON 生成后，您可以找到“复制到剪贴板”按钮，点击该按钮，JSON 将被复制到您的剪贴板。
* `保存到文件`。（`Excel for Mac` 用户不可用）JSON 生成后，您可以找到“另存为”按钮，点击该按钮，系统将提示您将 JSON 保存到文件。




## 3.3 转换设置

    Excel 单元格中的换行将呈现为 `\n`


### 选择标题、行/列
默认情况下，Excel 转 JSON 将第一行作为标题，但您也可以选择“第一列作为标题”。

如果选择 `第一行作为标题`：

第一（左）列将被视为标题列（JSON 对象的键），右侧的列将成为 JSON 对象的值。"

如果选择 `第一列作为标题`：

第一（上）行将被视为标题行（JSON 对象的键），下方的行将成为 JSON 对象的值。


### 转换模式
* `扁平 JSON 模式`
    * 简单地将 Excel 数据表转换为扁平 JSON。
* `嵌套 JSON 模式`
    * 首先将 Excel 数据表转换为扁平 JSON
    * 然后，使用“Flat”[https://www.npmjs.com/package/flat](https://www.npmjs.com/package/flat) 将带有分隔键的对象展开
    * Excel 转 JSON 调用 unflatten() 方法，分隔符为“.”，overwrite 为 true。如果您已订阅“专业功能”，则可以设置其他分隔符。

### 嵌套 JSON 键分隔符

用于分隔 JSON 输出中嵌套属性的分隔符。

- `点 (.)`
- `下划线 (_)` [专业功能]
- `双下划线 (__)` [专业功能]
- `斜杠 (/)` [专业功能]

### 其他设置

其他设置将在 [专业功能](profeatures.md) 中描述

# 4. 示例


## 4.1 示例 Excel 表格 1


|姓名|年龄|公司|
|---|---|---|
|David|27|WTSolutions|
|Ton|26|WTSolutions|
|Kitty|30|Microsoft|
|Linda|30|Microsoft|
|Joe|40|Github|

> 使用扁平 JSON 模式

**示例 JSON**

```json
[
    {
        "姓名": "David",
        "年龄": 27,
        "公司": "WTSolutions"
    },
    {
        "姓名": "Ton",
        "年龄": 26,
        "公司": "WTSolutions"
    },
    {
        "姓名": "Kitty",
        "年龄": 30,
        "公司": "Microsoft"
    },
    {
        "姓名": "Linda",
        "年龄": 30,
        "公司": "Microsoft"
    },
    {
        "姓名": "Joe",
        "年龄": 40,
        "公司": "Github"
    }
]
```

## 4.2 示例 Excel 表格 2

|id|student.name|student.familyname|student.age|
|---|---|---|---|
|1|Meimei|Han|12|
|2|Lily|Jaskson|15|
|3|Elon|Mask|18|

> 使用扁平 JSON 模式

```json
[{
	"id": 1,
	"student.name": "Meimei",
	"student.familyname": "Han",
	"student.age": 12
}, {
	"id": 2,
	"student.name": "Lily",
	"student.familyname": "Jaskson",
	"student.age": 15
}, {
	"id": 3,
	"student.name": "Elon",
	"student.familyname": "Mask",
	"student.age": 18
}]
```

> 使用嵌套 JSON 模式，并以点 . 作为分隔符

```json
[{
	"id": 1,
	"student": {
		"name": "Meimei",
		"familyname": "Han",
		"age": 12
	}
}, {
	"id": 2,
	"student": {
		"name": "Lily",
		"familyname": "Jaskson",
		"age": 15
	}
}, {
	"id": 3,
	"student": {
		"name": "Elon",
		"familyname": "Mask",
		"age": 18
	}
}]

```
