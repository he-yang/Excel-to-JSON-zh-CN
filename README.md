# Excel-to-JSON

[![Documentation Status](https://readthedocs.org/projects/excel-to-json/badge/?version=latest)](http://excel-to-json.readthedocs.io/en/latest/?badge=latest)

将Excel转换为JSON的插件

Excel-to-JSON现已在Microsoft Appsource(原Office商店)上架。 https://store.office.com/app.aspx?assetid=WA104380263

## 文档
[https://excel-to-json.wtsolutions.cn/zh-cn/latest/](https://excel-to-json.wtsolutions.cn/zh-cn/latest/)

## 开始使用

Excel转JSON是一款Microsoft Excel插件，可将Excel数据转换为JSON格式。

## 系统要求
本插件支持Excel 2013(或更高版本)、Excel Online、Office 365和Mac版Excel。

## 获取插件
* 在Excel 2013/2016或更高版本、Excel Online或Office 365中打开新工作表
* 选择**插入**选项卡或**开始**选项卡>加载项
* 在加载项搜索框中搜索"Excel-to-JSON"
* 点击加载项启动它
* 您将在Excel的**开始**选项卡中看到"Excel-to-JSON"按钮，现在可以开始使用此插件

### 使用插件

请注意，您需要选择至少两行数据，第一行将被视为表头。

* 准备您的Excel表格
* 选择要转换的数据
* 选择模式：Flat或Nested JSON模式
* 如果您已订阅"专业功能"，可以设置更多选项
* 点击"开始"按钮
* 稍后您将在"开始"按钮下方看到转换后的JSON
* 您可以"复制粘贴"或"复制到剪贴板"JSON并保存到电脑

## 示例

**示例Excel表格1**


|姓名|年龄|公司|
|---|---|---|
|David|27|WTSolutions|
|Ton|26|WTSolutions|
|Kitty|30|Microsoft|
|Linda|30|Microsoft|
|Joe|40|Github|

> 使用Flat JSON模式

**示例JSON**

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

**示例Excel表格2**

|ID|学生.姓名|学生.姓氏|学生.年龄|
|---|---|---|---|
|1|Meimei|Han|12|
|2|Lily|Jaskson|15|
|3|Elon|Mask|18|

> 使用Flat JSON模式

```json
[{
	"ID": 1,
	"学生.姓名": "Meimei",
	"学生.姓氏": "Han",
	"学生.年龄": 12
}, {
	"ID": 2,
	"学生.姓名": "Lily",
	"学生.姓氏": "Jaskson",
	"学生.年龄": 15
}, {
	"ID": 3,
	"学生.姓名": "Elon",
	"学生.姓氏": "Mask",
	"学生.年龄": 18
}]
```

> 使用Nested JSON模式

```json
[{
	"ID": 1,
	"学生": {
		"姓名": "Meimei",
		"姓氏": "Han",
		"年龄": 12
	}
}, {
	"ID": 2,
	"学生": {
		"姓名": "Lily",
		"姓氏": "Jaskson",
		"年龄": 15
	}
}, {
	"ID": 3,
	"学生": {
		"姓名": "Elon",
		"姓氏": "Mask",
		"年龄": 18
	}
}]
```
