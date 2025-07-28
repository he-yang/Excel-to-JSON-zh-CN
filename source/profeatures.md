# 6. Excel数据和转换设置

[English](https://excel-to-json.wtsolutions.cn/en/latest/profeatures.html)

WTSolutions的Excel转JSON工具提供了一系列增强工具功能的转换规则。
这些标记为[Pro Feature](pricing.md)的规则仅对已订阅工具的用户可用。

## 6.1 Excel数据

### Excel数据源
||[Web应用](WebApp.md)|[Excel加载项](ExcelAddIn.md)|[API](API.md)|[MCP](MCP.md)|
|:--:|:--:|:--:|:--:|:--:|
|适用|✅|✅|❌|❌|

您可以通过两种方式将Excel数据输入到Excel转JSON Web应用和Excel加载项：

* `在网页浏览器中加载Excel到JSON`
    * 复制并粘贴您的Excel数据到文本区域
    * 您可以从Excel、Google表格或任何其他兼容Excel的软件中复制和粘贴Excel数据，数据以制表符分隔
    * 您也可以复制和粘贴逗号分隔的CSV数据
* `在Excel中侧载Excel到JSON`：
    * 使用鼠标直接从Excel工作表中选择数据。
    * 或者，让Excel转JSON转换Excel中所有可见工作表。[Pro Feature](pricing.md)
    * 或者，让Excel转JSON转换Excel中的所有工作表。[Pro Feature](pricing.md)


## 6.2 转换设置

### 选择标题行或列
||[Web应用](WebApp.md)|[Excel加载项](ExcelAddIn.md)|[API](API.md)|[MCP](MCP.md)|
|:--:|:--:|:--:|:--:|:--:|
|适用|✅|✅|✅|❌|

您可以指定：
- 第一行作为标题
    - 第一行将被视为"标题"行，并将用作列名，后续行将被视为JSON值。
    - 后续行将被视为"数据"行，并将被视为JSON值。
- 第一列作为标题 [Pro Feature](pricing.md)
    - 第一列将被视为"标题"列，并将用作JSON键，后续列将被视为JSON值。
    - 右侧的列将被视为"数据"列，并将被视为JSON值。

### 转换模式
||[Web应用](WebApp.md)|[Excel加载项](ExcelAddIn.md)|[API](API.md)|[MCP](MCP.md)|
|:--:|:--:|:--:|:--:|:--:|
|适用|✅|✅|✅|❌|

转换模式指定如何将Excel数据转换为JSON。默认模式是`Flat`。

* `flat`：将Excel数据转换为扁平结构的JSON。
* `nested`：将Excel数据转换为嵌套结构的JSON。[Pro Feature](pricing.md)

### 嵌套JSON键分隔符
||[Web应用](WebApp.md)|[Excel加载项](ExcelAddIn.md)|[API](API.md)|[MCP](MCP.md)|
|:--:|:--:|:--:|:--:|:--:|
|适用|✅|✅|✅|❌|

嵌套JSON键分隔符指定在嵌套JSON模式中如何分隔嵌套键。默认分隔符是点(.)。

您还可以使用其他分隔符，如：
- `点(.)`或`"。"`
- `下划线(_)`或`"_"` [Pro Feature](pricing.md)
- `双下划线(__)`或`"__"` [Pro Feature](pricing.md)
- `正斜杠(/)`或`"/"` [Pro Feature](pricing.md)

例如，使用下划线(_)作为分隔符：

|id|<mark>student_name</mark>|student_familyname|student_age|
|---|---|---|---|
|1|Meimei|Han|12|
|2|Lily|Jaskson|15|

使用点(.)作为分隔符：
|id|<mark>student.name</mark>|student.familyname|student.age|
|---|---|---|---|
|1|Meimei|Han|12|
|2|Lily|Jaskson|15|

使用正斜杠(/)作为分隔符：
|id|<mark>student/name</mark>|student/familyname|student/age|
|---|---|---|---|
|1|Meimei|Han|12|
|2|Lily|Jaskson|15|

将生成：

> 使用嵌套JSON模式，分别使用分隔符"_"、"."、"/"。

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
}]
```
### 空单元格的输出格式
||[Web应用](WebApp.md)|[Excel加载项](ExcelAddIn.md)|[API](API.md)|[MCP](MCP.md)|
|:--:|:--:|:--:|:--:|:--:|
|适用|✅|✅|✅|❌|

空单元格选项通过三种方式处理Excel中的空单元格：

1. `空字符串 ""` 或 `"emptyString"`：将空单元格转换为空字符串 `""`
2. `JSON null` 或 `"null"`：将空单元格转换为 `null` [Pro Feature](pricing.md)
3. `不包含在JSON中` 或 `"exclude"`：从JSON中排除空单元格 [Pro Feature](pricing.md)。此选项在2DArray格式输出JSON中不可用。

**示例Excel表格**

|Name|Age|Company|
|---|---|---|
|David|27|WTSolutions|
|Ton||Microsoft|

> 使用空单元格 = "null"

```json
[{
    "Name": "David",
    "Age": 27,
    "Company": "WTSolutions"
}, {
    "Name": "Ton",
    // 看这里
    "Age": null,
    "Company": "Microsoft"
}]
```
> 使用空单元格 = "空字符串"
```json
[{
    "Name": "David",
    "Age": 27,
    "Company": "WTSolutions"
}, {
    "Name": "Ton",
    // 看这里
    "Age": "",
    "Company": "Microsoft"
}]
```
> 使用空单元格 = "不包含在JSON中"
```json
    "Name": "David",
    "Age": 27,
    "Company": "WTSolutions"
}, {
    "Name": "Ton",
    // 看这里
    "Company": "Microsoft"
}]
```

### 布尔值的输出格式
||[Web应用](WebApp.md)|[Excel加载项](ExcelAddIn.md)|[API](API.md)|[MCP](MCP.md)|
|:--:|:--:|:--:|:--:|:--:|
|适用|✅|✅|✅|❌|

布尔格式指定如何转换Excel中的布尔值：

1. `JSON true/false` 或 `"trueFalse"`：转换为JSON布尔值 (`true` 或 `false`)
2. `字符串 "true"/"false"` 或 `"string"`：转换为字符串 (`"true"` 或 `"false"`) [Pro Feature](pricing.md)
3. `数字 1/0` 或 `"10"`：转换为数字 (`1` 表示 `TRUE`，`0` 表示 `FALSE`) [Pro Feature](pricing.md)

**示例Excel表格**

|Name|IsStudent|IsEmployed|
|---|---|---|
|David|TRUE|FALSE|
|Ton|FALSE|TRUE|

> 使用布尔格式 = "JSON true/false"

```json
[{
    "Name": "David",
    // 看这里
    "IsStudent": true,
    "IsEmployed": false
}, {
    "Name": "Ton",
    "IsStudent": false,
    "IsEmployed": true
}]
```
> 使用布尔格式 = "字符串 \"true\"/\"false\""
```json
[{
    "Name": "David",
    // 看这里
    "IsStudent": "true",
    "IsEmployed": "false"
}, {
    "Name": "Ton",
    "IsStudent": "false",
    "IsEmployed": "true"
}]
```
> 使用布尔格式 = "数字 1/0"
```json
[{
    "Name": "David",
    // 看这里
    "IsStudent": 1,
    "IsEmployed": 0
}, {
    "Name": "Ton",
    "IsStudent": 0,
    "IsEmployed": 1
}]
```

### 日期格式值的输出格式
||[Web应用](WebApp.md)|[Excel加载项](ExcelAddIn.md)|[API](API.md)|[MCP](MCP.md)|
|:--:|:--:|:--:|:--:|:--:|
|适用|❌|✅|❌|❌|

此功能允许您将Excel中的日期值转换为特定格式，如ISO 8601或自1900-01-01以来的天数。当您在网页浏览器中加载Excel到JSON时，此功能不可用。如果您想将Excel中的日期值转换为特定格式，可以在Excel应用程序中侧载Excel到JSON。

日期格式指定如何转换Excel中的日期值：

1. `自1900-01-01以来的天数`：转换为自1900-01-01以来的天数
2. `字符串，ISO 8601 (YYYY-MM-DDTHH:mm:ssZ)`：转换为ISO 8601格式的字符串 [Pro Feature](pricing.md)


> <mark> 注意：日期格式功能仅在您在Excel数据表第一行添加 $date$ 作为后缀时才有效。请参考后续示例标题行。 </mark>
> 注意：Excel无法将1900-01-01之前的日期渲染为日期格式，因此您可能会发现该单元格被解释为字符串格式。
> 例如，如果您想将Birthday列转换为ISO8601格式，您应该在列标题中添加$date$作为后缀，如下例所示。

**示例Excel表格**

|Name|<mark>Birthday$date$</mark>|
|---|---|
|David|1900-01-01|
|Ton|1995-05-15|

> 使用日期格式 = "自1900-01-01以来的天数"

```json
[{
    "Name": "David",
    "Birthday": 1
}, {
    "Name": "Ton",
    // 看这里
    "Birthday": 34834
}]
```
> 使用日期格式 = "字符串，ISO 8601 (YYYY-MM-DDTHH:mm:ssZ)"
```json
[{
    "Name": "David",
    "Birthday": "1900-01-01T00:00:00.000Z"
}, {
    "Name": "Ton",
    // 看这里
    "Birthday": "1995-05-15T00:00:00.000Z"
}]
```

### JSON的输出格式
||[Web应用](WebApp.md)|[Excel加载项](ExcelAddIn.md)|[API](API.md)|[MCP](MCP.md)|
|:--:|:--:|:--:|:--:|:--:|
|适用|✅|✅|✅|❌|

您可以通过两种方式呈现输出JSON：
1. `对象数组` 或 `"arrayOfObject"`
2. `二维数组` 或 `"2DArray"` [Pro Feature](pricing.md)

**示例Excel表格**

|Name|Age|Company|
|---|---|---|
|David|27|WTSolutions|
|Ton|25|Microsoft|

> 使用"对象数组"选项。

```json
[{
    "Name": "David",
    "Age": 27,
    "Company": "WTSolutions"        
},
{
    "Name": "Ton",
    "Age": 25,
    "Company": "Microsoft"  
}
]
```
> 使用"二维数组"选项。
```json
[
    ["Name", "Age", "Company"]
    ["David", 27, "WTSolutions"],
    ["Ton", 25, "Microsoft"]
]
```

### 单对象JSON的输出格式
||[Web应用](WebApp.md)|[Excel加载项](ExcelAddIn.md)|[API](API.md)|[MCP](MCP.md)|
|:--:|:--:|:--:|:--:|:--:|
|适用|✅|✅|✅|❌|

当转换后的JSON中只有一个对象时，此选项指定如何保留单对象JSON：

1. `数组 []` 或 `"array"`：转换为包含一个对象的数组
2. `对象 {}` 或 `"object"`：转换为一个对象 [Pro Feature](pricing.md)


> 注意：此功能仅在转换后的JSON中只有一个对象且JSON的输出格式为对象数组时才有效。

**示例Excel表格**

|Name|Age|
|---|---|
|David|20|

> 使用"数组 []"

```json
[{
    "Name": "David",
    "Age": 20
}]
```
> 使用"对象 {}"
```json
{
    "Name": "David",
    "Age": 20
}
```

### 另存为时的文件名
||[Web应用](WebApp.md)|[Excel加载项](ExcelAddIn.md)|[API](API.md)|[MCP](MCP.md)|
|:--:|:--:|:--:|:--:|:--:|
|适用|✅|✅|❌|❌|

另存为文件名功能允许您为JSON输出文件指定自定义文件名，当您在转换后点击"另存为"按钮时。当您在此字段中输入文件名时，转换后的JSON数据将以您指定的名称保存，而不是默认名称(excel-to-json.json)。

**文件名要求：**
- 最大长度：200个字符（不包括扩展名）
- 文件扩展名将自动设置为.json
- 不能以点(.)或空格开头或结尾
- 不能使用Windows保留名称（例如，CON、PRN、AUX等）
- 不能包含以下字符：< > : \ / | ? *

**示例：**
- 有效文件名：data.json、my_data.json、export_2024.json
- 无效文件名：.data.json、con.json、my:data.json

**支持的平台：**
- Windows上的Excel：是
- Office.com或Onedrive上的Excel网页版：是
- Mac版Excel：不支持
  - 如果您是Mac用户，可以使用Office.com或Onedrive上的Excel网页版。


## 6.3 更多功能

如果您已订阅并希望看到更多功能，请通过电子邮件he.yang@wtsolutions.cn联系我们。

## 6.4 Pro Code

Pro Code是您在Stripe上购买Excel-to-JSON加载项时使用的`电子邮件地址`。此代码是访问专业功能所必需的。

## 6.5 售后服务

如有任何问题或疑虑，您可以通过电子邮件he.yang@wtsolutions.cn联系我们。我们将尽力在24小时内回复您，但不会超过72小时。如果您的问题与订阅相关，请在电子邮件中包含您的`Pro Code`。
