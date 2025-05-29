# 使用说明

[English](https://excel-to-json.wtsolutions.cn/en/latest/usage.html)

本文档适用于Excel-to-JSON版本2.3.0

强烈建议先阅读[入门指南](getstarted.md)部分。

    Excel单元格中的换行符将被渲染为`\n`

<a name="Conversiontypes"></a>
## 转换

注意：您需要至少选择两行数据，第一行将被视为表头。

* 本插件会首先收集您**选定**的所有单元格
* 第一行将被视为表头
* 后续行将与表头进行映射，如下例所示

## 转换模式
* Flat JSON模式
    * 简单地将Excel数据表转换为Flat的JSON格式
* Nested JSON模式
    * 首先将Excel数据表转换为Flat JSON
    * 然后使用"Flat"包[https://www.npmjs.com/package/flat](https://www.npmjs.com/package/flat)对带有分隔符的键进行解平铺操作
    * Excel-to-JSON调用unflatten()函数时，分隔符默认为"."，overwrite参数为true。如果您订阅了"专业功能"，可以设置其他分隔符

<script async src="https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=ca-pub-8772217510669640"
     crossorigin="anonymous"></script>
<ins class="adsbygoogle"
     style="display:block; text-align:center;"
     data-ad-layout="in-article"
     data-ad-format="fluid"
     data-ad-client="ca-pub-8772217510669640"
     data-ad-slot="2653271427"></ins>
<script>
     (adsbygoogle = window.adsbygoogle || []).push({});
</script>

### 使用插件的视频演示

<iframe src="//player.bilibili.com/player.html?isOutside=true&aid=114345920959097&bvid=BV1jdoAYHEDF&cid=29442509957&p=1" scrolling="no" border="0" frameborder="no" framespacing="0" allowfullscreen="true"></iframe>

## 示例

**Excel表格示例1**

|姓名|年龄|公司|
|---|---|---|
|David|27|WTSolutions|
|Ton|26|WTSolutions|
|Kitty|30|Microsoft|
|Linda|30|Microsoft|
|Joe|40|Github|

> 使用Flat JSON模式

**JSON示例**

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

**Excel表格示例2**

|id|学生.姓名|学生.姓氏|学生.年龄|
|---|---|---|---|
|1|Meimei|Han|12|
|2|Lily|Jaskson|15|
|3|Elon|Mask|18|

> 使用Flat JSON模式

```json
[{
	"id": 1,
	"学生.姓名": "Meimei",
	"学生.姓氏": "Han",
	"学生.年龄": 12
}, {
	"id": 2,
	"学生.姓名": "Lily",
	"学生.姓氏": "Jaskson",
	"学生.年龄": 15
}, {
	"id": 3,
	"学生.姓名": "Elon",
	"学生.姓氏": "Mask",
	"学生.年龄": 18
}]
```

> 使用Nested JSON模式

```json
[{
	"id": 1,
	"学生": {
		"姓名": "Meimei",
		"姓氏": "Han",
		"年龄": 12
	}
}, {
	"id": 2,
	"学生": {
		"姓名": "Lily",
		"姓氏": "Jaskson",
		"年龄": 15
	}
}, {
	"id": 3,
	"学生": {
		"姓名": "Elon",
		"姓氏": "Mask",
		"年龄": 18
	}
}]
```

<a name="jsonOutput"></a>
### JSON输出

有几种方式可以将生成的JSON保存到本地计算机：

* 复制粘贴。JSON生成后，您可以在插件中看到它们，并可以简单地复制粘贴到任何地方
* Copy to Clipboard。JSON生成后，您可以找到"Copy to Clipboard"按钮，点击该按钮，JSON将被复制到剪贴板
* Save As （Mac用户不可用）。JSON生成后，您可以找到"Save As"按钮，点击该按钮，您可以将JSON保存到本地计算机

## 专业功能
Excel-to-JSON为有效订阅用户提供专业功能：

* **嵌套分隔符**：支持自定义嵌套JSON键的分隔符(/, _, .)
* **空单元格处理**：空单元格的三种处理方式：空字符串""、JSON null或直接排除
* **布尔值格式**：三种布尔值转换格式：JSON true/false、字符串"true"/"false"或数字1/0
* **单对象json输出格式**：当只有一行数据时，是否输出为单对象JSON格式，或者输出为数组格式
* **日期格式**：日期转换选项：自1990-01-01以来的天数或ISO 8601字符串格式
* **另存为文件名**: 在使用Save As功能时，您可以自定义文件名

专业功能的详细信息请参阅[专业功能](profeatures.md)部分。
