# 4. API (Excel to JSON by HTTPS POST request)

[English](https://excel-to-json.wtsolutions.cn/en/latest/API.html)

WTSolutions的Excel转JSON是一系列工具，可将Excel转换为JSON，支持转换为扁平结构和嵌套结构JSON。它提供了"Excel转JSON"的全场景解决方案，包括Excel加载项、Web应用、开放API和企业级MCP工具：

* [Web应用：直接在Web浏览器中转换Excel到JSON。](WebApp.md)
* [Excel加载项：在Excel中转换Excel到JSON，与Excel环境无缝集成。](ExcelAddIn.md)
* <mark>API：通过HTTPS POST请求转换Excel到JSON。</mark> (<-- 您当前所在位置。)
* [MCP服务：通过AI模型MCP的SSE/流式HTTP请求转换Excel到JSON。](MCP.md)

## 4.1 需求

需要HTTPS POST请求工具，如Postman、Curl、Python Requests、Javascript fetch等。智能体平台如coze。
确保通过设置CORS头来正确处理跨域问题。

## 4.2 访问方式

向访问点 `https://mcp.wtsolutions.cn/excel-to-json-api` 发送`POST`请求，并包含使用部分中描述的必要参数。使用此API有两种方式：

* 标准方式（第4.3节）：免费使用，采用标准转换规则。
* 专业方式（第4.4节）：支持自定义转换规则，需要有效的[WTSolutions Excel转JSON服务订阅](pricing.md)。

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

## 4.3 使用方法 -- 标准方式

Excel转JSON API提供了一种简单的方式将Excel和CSV数据转换为JSON格式。此API接受：
- 制表符分隔或逗号分隔的文本数据，或
- 指向Excel文件的URL

在本节中，您可以找到使用此API的标准方式，这种方式是免费的。如果您想进行自定义转换，请参阅第4.4节 使用方法 -- 专业方式。

### 4.3.1 请求格式

API接受包含以下参数之一的`application/json`格式的POST请求：

| 参数 | 类型 | 是否必填 | 描述 |
|------|------|----------|------|
| data | 字符串 | 否 | 制表符分隔或逗号分隔的文本数据，至少包含两行（标题行+数据行）。必须提供'data'或'url'中的一个 |
| url | 字符串 | 否 | 指向Excel或CSV文件的URL。必须提供'data'或'url'中的一个 |

> 注意：
> - 请提供`data`或`url`中的一个，不要同时提供。

#### 对data和url的要求

发送`data`时：
- 输入数据必须是制表符分隔（Excel）或逗号分隔（CSV）的文本，至少包含两行（标题行+数据行）。
  1. 第一行将被视为"标题"行，API将使用它作为列名，即JSON键。
  2. 后续行将被视为"数据"行，API将把它们视为JSON值。

发送`url`时：
- Excel文件的每个工作表应至少包含两行（标题行+数据行）。
  1. 第一行将被视为"标题"行，API将使用它作为列名，即JSON键。
  2. 后续行将被视为"数据"行，API将把它们视为JSON值。
- 此Excel文件应为'.xlsx'格式。
- Excel文件的每个工作表将被转换为一个JSON对象。
- 每个JSON对象将具有'sheetName'（字符串）和'data'（对象数组）属性。
- 'data'数组中的每个JSON对象将具有与列名对应的属性。
- 'data'数组中的每个JSON对象将具有与单元格值对应的值。



### 4.3.2 响应格式
API返回具有以下结构的JSON对象：

| 字段 | 类型 | 描述 |
|------|------|------|
| isError | 布尔值 | 指示处理请求时是否出错 |
| msg | 字符串 | 'success'或错误描述 |
| data | 字符串 | 如果使用URL，则为工作表对象数组；如果使用直接数据，则为字符串 |


### 4.3.3 示例

#### 使用'data'的请求示例

原始Excel数据：

| Name       | Age | IsStudent |
|------------|-----|-----------|
| John Doe   | 25  | false     |
| Jane Smith | 30  | true      |

请求：

```json
{"data": "Name\tAge\tIsStudent\nJohn Doe\t25\tfalse\nJane Smith\t30\ttrue"}
```

响应：

```json
{
	"isError": false,
	"msg": "success",
	"data": "[{\"Name\":\"John Doe\",\"Age\":25,\"IsStudent\":false},{\"Name\":\"Jane Smith\",\"Age\":30,\"IsStudent\":true}]"
}
```

#### 使用'url'的请求示例

`Sheet1`中的原始Excel数据：

| Name       | Age | IsStudent |
|------------|-----|-----------|
| John Doe   | 25  | false     |
| Jane Smith | 30  | true      |

`Sheet2`中的原始Excel数据：

| ID | Value   |
|----|---------|
| 1  | Example |

请求：

```json
{"url": "https://tools.wtsolutions.cn/example.xlsx"}
```

响应：

```json
{
	"isError": false,
	"msg": "success",
	"data": "[{\"sheetName\":\"Sheet1\",\"data\":[{\"Name\":\"John Doe\",\"Age\":25,\"IsStudent\":false},{\"Name\":\"Jane Smith\",\"Age\":30,\"IsStudent\":true}]},{\"sheetName\":\"Sheet2\",\"data\":[{\"ID\":1,\"Value\":\"Example\"}]}]"
}
```

#### 错误响应示例
```json
{
     "isError": true,
     "msg": "Excel数据至少需要2行",
     "data": ""
}
```

### 4.3.4 数据类型处理
此API的标准方式会自动检测并转换不同的数据类型，而在专业方式中，用户可以通过在请求体中提供'options'对象来自定义转换规则，如第4.4节所述。
- **数字**：转换为数值
- **布尔值**：识别'true'/'false'（不区分大小写）并转换为布尔值
- **日期**：检测各种日期格式并适当转换
- **字符串**：视为字符串值
- **空值**：表示为空字符串


## 4.4 使用方法 -- 专业方式

本节 -- 专业方式适用于已购买[Excel转JSON服务订阅](pricing.md)的用户。如果您尚未购买订阅，请参阅第4.3节 使用方法 -- 标准方式。


### 4.4.1 请求格式

API接受包含以下参数之一的`application/json`格式的POST请求：

| 参数 | 类型 | 是否必填 | 描述 |
|------|------|----------|------|
| data | 字符串 | 否 | 制表符分隔或逗号分隔的文本数据，至少包含两行（标题行+数据行）。必须提供'data'或'url'中的一个 |
| url | 字符串 | 否 | 指向Excel或CSV文件的URL。必须提供'data'或'url'中的一个 |
| options | 对象 | 是 | 用于自定义转换过程的配置对象 |

> 注意：
> - 请提供`data`或`url`中的一个，不要同时提供。
> - 如果您想使用自定义转换设置，则`options`是必需的。如果您没有有效的专业代码，请参阅第4.3节 使用方法 -- 标准方式。

#### 对data和url的要求

发送`data`时：
- 输入数据必须是制表符分隔（Excel）或逗号分隔（CSV）的文本，至少包含两行（标题行+数据行）或两列（标题列+数据列），具体取决于options参数中的"header"设置。
  1. 第一行或第一列将被视为"标题"行/列，API将使用它作为列名，即JSON键。
  2. 后续行或列将被视为"数据"行/列，API将把它们视为JSON值。

发送`url`时：
- Excel文件的每个工作表应至少包含两行或两列（标题行/列+数据行/列）。
  1. 第一行或第一列将被视为"标题"行/列，API将使用它作为列名，即JSON键。
  2. 后续行或列将被视为"数据"行/列，API将把它们视为JSON值。
- 此Excel文件应为'.xlsx'格式。
- Excel文件的每个工作表将被转换为一个JSON对象。
- 每个JSON对象将具有'sheetName'（字符串）和'data'属性。
- 'data'数组中的每个JSON对象将具有与标题名称对应的属性。
- 'data'数组中的每个JSON对象将具有与单元格值对应的值。


### 4.4.2 选项对象
可选的`options`对象可以包含以下属性：

| 属性 | 类型 | 默认值 | 描述 |
|------|------|--------|------|
| [proCode](profeatures.md#pro-code) | 字符串 | "" | 用于自定义转换规则的专业代码，需要有效的[Excel转JSON服务订阅](pricing.md)。proCode可以通过options中参数传递，也可以在https请求头中传递proCode。 |
| [jsonMode](profeatures.md#id5) | 字符串 | "flat" | JSON输出的格式模式："nested"（嵌套）或"flat"（扁平） |
| [header](profeatures.md#id4) | 字符串 | "row" | 指定使用哪一行/列作为标题："row"（第一行）或"column"（第一列） |
| [delimiter](profeatures.md#json) | 字符串 | "." | 当使用`jsonMode`为"nested"时，嵌套JSON键的分隔符，可接受的分隔符有"."、"_"、"__"、"/"。 |
| [emptyCell](profeatures.md#id6) | 字符串 | "emptyString" | 空单元格的处理方式："emptyString"（空字符串）、"null"（空值）或"exclude"（排除） |
| [booleanFormat](profeatures.md#id7) | 字符串 | "trueFalse" | 布尔值的格式："trueFalse"（true/false）、"10"（1/0）或"string"（字符串） |
| [jsonFormat](profeatures.md#id9) | 字符串 | "arrayOfObject" | JSON输出的整体格式："arrayOfObject"（对象数组）或"2DArray"（二维数组） |
| [singleObjectFormat](profeatures.md#id10) | 字符串 | "array" | 当结果只有一个对象时的格式："array"（保持为数组）或"object"（返回为单个对象） |

> 注意：
> - `proCode`可以在options中参数传递，也可以在https请求头中传递proCode。`proCode`是必须的，如果您没有有效的专业代码，请参阅第4.3节 使用方法 -- 标准方式。
> - `delimiter`仅在`jsonMode`为"nested"时有效。
> - `singleObjectFormat`仅在`jsonFormat`为"arrayOfObject"时有效。
> - `jsonFormat`为"2DArray"仅在`jsonMode`为"flat"时有效。
> - `proCode`是必需的。如果您没有有效的专业代码，请参阅第4.3节 使用方法 -- 标准方式。
> - 详细的转换规则可以在[专业功能](profeatures.md)中找到。

### 4.4.3 响应格式
API返回具有以下结构的JSON对象：

| 字段 | 类型 | 描述 |
|------|------|------|
| isError | 布尔值 | 指示处理请求时是否出错 |
| msg | 字符串 | 'success'或错误描述 |
| data | 字符串 | 如果使用URL，则为工作表对象数组；如果使用直接数据，则为字符串；如果出错，则为空字符串。|


### 4.4.4 示例

#### 使用'data'的请求示例

原始数据：

| id | student.name | student.familyname | student.age |
|----|--------------|--------------------|-------------|
| 1  | Meimei       | Han                | 12          |
| 2  | Lily         | Jaskson            | 15          |
| 3  | Elon         | Mask               | 18          |

请求1：

```json
{
     "data": "id\tstudent.name\tstudent.familyname\tstudent.age\n1\tMeimei\tHan\t12\n2\tLily\tJaskson\t15\n3\tElon\tMask\t18",
     "options": {
         "proCode": "proCode(please input your proCode)",
         "jsonMode": "nested",
         "delimiter": "."
     }
}
```
响应1：
```json
{
	"isError": false,
	"msg": "success",
	"data": "[{\"id\":1,\"student\":{\"name\":\"Meimei\",\"familyname\":\"Han\",\"age\":12}},{\"id\":2,\"student\":{\"name\":\"Lily\",\"familyname\":\"Jaskson\",\"age\":15}},{\"id\":3,\"student\":{\"name\":\"Elon\",\"familyname\":\"Mask\",\"age\":18}}]"
}
```

请求2：
```json
{
     "data": "id\tstudent.name\tstudent.familyname\tstudent.age\n1\tMeimei\tHan\t12\n2\tLily\tJaskson\t15\n3\tElon\tMask\t18",
     "options": {
         "proCode": "proCode(please input your proCode)",
         "jsonMode": "flat",
         "jsonFormat": "2DArray"
     }
}
```


响应2：
```json
{
	"isError": false,
	"msg": "success",
	"data": "[[\"id\",\"student.name\",\"student.familyname\",\"student.age\"],[1,\"Meimei\",\"Han\",12],[2,\"Lily\",\"Jaskson\",15],[3,\"Elon\",\"Mask\",18]]"
}
```

#### 使用'url'的成功响应示例（多工作表Excel）

`Sheet1`中的原始Excel数据：

| Name       | Age | IsStudent |
|------------|-----|-----------|
| John Doe   | 25  | false     |
| Jane Smith | 30  | true      |

`Sheet2`中的原始Excel数据：

| ID | Value   |
|----|---------|
| 1  | Example |


请求：

```json
{
     "url": "https://tools.wtsolutions.cn/example.xlsx",
     "options": {
         "proCode": "proCode(please input your proCode)",
         "booleanFormat": "10",
         "jsonFormat": "arrayOfObject",
         "singleObjectFormat": "object"
     }
}
```

响应：

```json
{
	"isError": false,
	"msg": "success",
	"data": "[{\"sheetName\":\"Sheet1\",\"data\":[{\"Name\":\"John Doe\",\"Age\":25,\"IsStudent\":0},{\"Name\":\"Jane Smith\",\"Age\":30,\"IsStudent\":1}]},{\"sheetName\":\"Sheet2\",\"data\":{\"ID\":1,\"Value\":\"Example\"}}]"
}
```

#### 错误响应示例
```json
{
     "isError": true,
     "msg": "Excel数据至少需要2行",
     "data": ""
}
```

## 4.5 错误处理
API针对常见问题返回描述性错误消息：
- `Excel Data Format Invalid`：当输入数据不是制表符分隔或逗号分隔时
- `At least 2 rows are required`：当输入数据少于2行时
- `Both data and url received`：当同时提供'data'和'url'参数时
- `Network Error when fetching file`：从提供的URL下载文件时出错
- `File not found`：找不到提供的URL上的文件
- `Blank/Null/Empty cells in the first row not allowed`：当标题行包含空单元格时
- `Server Internal Error`：发生意外错误时


## 4.6 使用视频

本视频展示的是在agents智能体平台coze上，利用本api创作的插件和智能体的使用。

<iframe src="//player.bilibili.com/player.html?isOutside=true&aid=114958255789128&bvid=BV13GhhzeEhJ&cid=31439980943&p=1" scrolling="no" border="0" frameborder="no" framespacing="0" allowfullscreen="true"></iframe>