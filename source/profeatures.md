# 5. 专业功能 Pro Features

[English](https://excel-to-json.wtsolutions.cn/latest/profeatures.html)

Excel转JSON加载项提供了一组专业功能，增强了加载项的功能性。这些功能仅对已订阅加载项的用户开放。

## 5.1 订阅、付款和取消

7天免费试用，之后您将按以下费率之一每月支付专业功能费用。您可以在第7天前随时取消订阅，不会被收费：
- 美元 US$2.66 / 月，
- 人民币 ¥19.90 / 月，
- 欧元 €2.36 / 月，
- 港币 HK$21.80 / 月
- 您的当地货币（如果Stripe支持）。

每个专业代码可支持10台设备访问专业功能。7天试用期结束后，您可以随时取消订阅，这将在当前 billing周期结束时生效。

每个专业代码对WTSolutions提供的Excel转JSON和[JSON转Excel](https://json-to-excel.wtsolutions.cn/en/latest/)加载项均有效。

### 通过Stripe订阅

[https://buy.stripe.com/00gdQT2iz0Vp32E002](https://buy.stripe.com/00gdQT2iz0Vp32E002)

支付方式：
- 银行卡（Visa、Mastercard、American Express、JCB、银联）
- Apple Pay（需要使用Apple/IOS/Mac设备打开上述链接）
- Google Pay（需要使用Google/Andriod设备打开上述链接）
- Alipay 支付宝
- Link

有关订阅条款，请参阅[使用条款](termsofuse.md)

### 取消订阅

您可以随时取消订阅。当前 billing周期结束后，您将不再有权访问专业功能。

[https://billing.stripe.com/p/login/5kQ4gyadZ7p22JY3V83Je00](https://billing.stripe.com/p/login/5kQ4gyadZ7p22JY3V83Je00)


## 5.2 专业功能 Pro Features

### 嵌套JSON键分隔符

嵌套JSON键分隔符指定在嵌套JSON模式下如何分隔嵌套键。默认分隔符是点(.)。

您还可以使用其他分隔符，如：
- `点 (.)`
- `下划线 (_)` [专业功能]
- `双下划线 (__)` [专业功能]
- `正斜杠 (/)` [专业功能]


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


> 分别使用"_", ".", "/"作为分隔符的嵌套JSON模式

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


空单元格选项提供了三种处理Excel中空单元格的方式：

1. `空字符串 ""`：将空单元格转换为空字符串`""`
2. `JSON null`：将空单元格转换为`null`
3. `不包含在JSON中`：从JSON中排除空单元格

**Excel示例表格**

|姓名|年龄|公司|
|---|---|---|
|David|27|WTSolutions|
|Ton||Microsoft|

> 使用 空单元格 = "null"

```json
[{
    "姓名": "David",
    "年龄": 27,
    "公司": "WTSolutions"
}, {
    "姓名": "Ton",
    // 注意这里
    "年龄": null,
    "公司": "Microsoft"
}]
```
> 使用 空单元格 = "空字符串"
```json
[{
    "姓名": "David",
    "年龄": 27,
    "公司": "WTSolutions"
}, {
    "姓名": "Ton",
    // 注意这里
    "年龄": "",
    "公司": "Microsoft"
}]
```
> 使用空单元格 = "不包含在JSON中"
```json
    "姓名": "David",
    "年龄": 27,
    "公司": "WTSolutions"
}, {
    "姓名": "Ton",
    // 注意这里
    "公司": "Microsoft"
}]
```

### 布尔值的输出格式

布尔格式指定如何转换Excel中的布尔值：

1. `JSON true/false`：转换为JSON布尔值（`true`或`false`）
2. `字符串 "true"/"false"`：转换为字符串（`"true"`或`"false"`）[专业功能]
3. `数字 1/0`：转换为数字（`1`表示`TRUE`，`0`表示`FALSE`）[专业功能]



**Excel示例表格**

|姓名|是学生|已就业|
|---|---|---|
|David|TRUE|FALSE|
|Ton|FALSE|TRUE|

> 使用 布尔值格式 = "JSON true/false"

```json
[{
    "姓名": "David",
    // 注意这里
    "是学生": true,
    "已就业": false
}, {
    "姓名": "Ton",
    "是学生": false,
    "已就业": true
}]
```
> 使用 布尔值格式 = "字符串 "true"/"false""
```json
[{
    "姓名": "David",
    // 注意这里
    "是学生": "true",
    "已就业": "false"
}, {
    "姓名": "Ton",
    "是学生": "false",
    "已就业": "true"
}]
```
> 使用 布尔值格式 = "数字 1/0"
```json
[{
    "姓名": "David",
    // 注意这里
    "是学生": 1,
    "已就业": 0
}, {
    "姓名": "Ton",
    "是学生": 0,
    "已就业": 1
}]
```

### 日期格式值的输出格式
此功能允许您将Excel中的日期值转换为特定格式，如ISO 8601或自1900-01-01以来的天数。当您在Web浏览器中加载Excel转JSON时，此功能不可用。如果您想将Excel中的日期值转换为特定格式，可以在Excel应用程序中旁加载Excel转JSON。

日期格式指定如何转换Excel中的日期值：

1. `自1900-01-01以来的天数`：转换为自1900-01-01以来的天数
2. `字符串，ISO 8601 (YYYY-MM-DDTHH:mm:ssZ)`：转换为ISO 8601格式的字符串 [专业功能]



> <mark>注意：日期格式功能仅在Excel数据表表头添加$date$后缀时才有效。</mark>
> 注意：Excel无法将1900-01-01之前的日期渲染为日期格式，因此这些单元格可能会被解释为字符串格式。
> 例如，如果您想将生日列转换为ISO8601格式，应在列标题中添加$date$后缀，如下例所示。

**Excel示例表格**

|姓名|<mark>生日$date$</mark>|
|---|---|
|David|1900-01-01|
|Ton|1995-05-15|

> 使用日期格式 = "从1900-01-01开始的天数"

```json
[{
    "姓名": "David",
    "生日": 1
}, {
    "姓名": "Ton",
    // 注意这里
    "生日": 34834
}]
```
> 使用日期格式 = "字符串, ISO 8601 (YYYY-MM-DDTHH:mm:ssZ)"
```json
[{
    "姓名": "David",
    "生日": "1900-01-01T00:00:00.000Z"
}, {
    "姓名": "Ton",
    // 注意这里
    "生日": "1995-05-15T00:00:00.000Z"
}]
```

### JSON的输出格式

您有两种选择来呈现输出JSON：
1. `对象数组`
2. `二维数组` [专业功能]


**例子Excel表**

|Name|Age|Company|
|---|---|---|
|David|27|WTSolutions|
|Ton|25|Microsoft|

> 选用 "Array of object" .

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
> 选用 "2D Array" option.
```json
[
    ["Name", "Age", "Company"]
    ["David", 27, "WTSolutions"],
    ["Ton", 25, "Microsoft"]
]
```

### 单个对象JSON的输出格式

当转换后的JSON中只有一个对象时，此选项指定如何保留单个对象JSON：

1. `数组 []`：转换为包含一个对象的数组
2. `对象 {}`：转换为一个对象 [专业功能]



**Excel示例表格**

|姓名|年龄|
|---|---|
|David|27|

> 使用单对象JSON输出格式 = "单对象JSON"

```json
{
    "姓名": "David",
    "年龄": 27
}
```
> 使用单对象JSON输出格式 = "数组JSON"
```json
[{
    "姓名": "David",
    "年龄": 27
}]
```



### 另存为时的文件名

另存为文件名功能允许您为JSON输出文件指定自定义文件名，当您在转换后点击"另存为"按钮时。当您在此字段中输入文件名时，转换后的JSON数据将以您指定的名称保存，而不是默认名称(excel-to-json.json)。



**文件名要求：**
- 最大长度：200个字符（不包括扩展名）
- 文件扩展名将自动设置为.json
- 不能以点(.)或空格开头或结尾
- 不能使用Windows保留名称（如CON、PRN、AUX等）
- 不能包含以下字符：< > : \ / | ? *

**示例：**
- 有效文件名：data.json、my_data.json、export_2024.json
- 无效文件名：.data.json、con.json、my:data.json

**支持平台：**
- Windows版Excel：支持
- Office.com或Onedrive上的Web版Excel：支持
- Mac版Excel：不支持
  - 如果您是Mac用户，可以使用Office.com或Onedrive上的Web版Excel


## 5.3 更多功能

如果您已订阅，并且希望看到更多功能，请发送电子邮件至he.yang@wtsolutions.cn

## 5.4 专业代码Pro Code

专业代码是您在Stripe上Excel转JSON加载项结账过程中使用的`电子邮件地址`。此代码是访问专业功能所必需的。

## 5.5 售后服务

如有任何问题或疑问，请通过邮件he.yang@wtsolutions.cn联系我们。我们将在24小时内尽快回复您，最迟不超过72小时。如果问题与您的订阅相关，请在邮件中包含您的`专业代码Pro Code`。
