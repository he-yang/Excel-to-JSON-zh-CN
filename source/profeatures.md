# 专业版功能 Pro Features

Excel-to-JSON插件提供了一系列专业版功能，这些功能仅对已订阅的用户开放。

## 订阅

7天免费试用，之后将按以下费率按月收费。您可以在第7天之前随时取消订阅，不会被扣费：
- 美元 US$2.66 / 月
- 人民币 ¥19.90 / 月
- 欧元 €2.36 / 月
- 港币 HK$21.80 / 月
- 如果Stripe支持您的本地货币，则按当地货币收费

每个专业代码Pro Code可支持10台设备访问专业版功能。7天试用期结束后，您可以随时取消订阅，取消将在当前计费周期结束时生效。

通过Stripe订阅：[https://buy.stripe.com/00gdQT2iz0Vp32E002](https://buy.stripe.com/00gdQT2iz0Vp32E002)
支付方式：
- 银行卡（Visa、Mastercard、American Express、JCB、银联）
- Apple Pay（需使用苹果设备打开上述链接）
- Google Pay（需使用安卓设备打开上述链接）
- Link

订阅条款请参考[使用条款](termsofuse.md)

![微信客服](https://invest.wtsolutions.cn/wechat.png)

## 专业代码

专业代码是您在Stripe上购买Excel-to-JSON插件时使用的`电子邮箱地址`。此代码是访问专业版功能的必要条件。

## 嵌套分隔符

嵌套分隔符指定了在嵌套JSON模式中如何分隔嵌套键。默认分隔符是点号(.)。

您也可以使用其他分隔符如：
- 斜杠(/)
- 下划线(_)
- 点号(.)

例如，使用下划线(_)作为分隔符：

|id|<mark>student_name</mark>|student_familyname|student_age|
|---|---|---|---|
|1|Meimei|Han|12|
|2|Lily|Jaskson|15|

使用点号(.)作为分隔符：
|id|<mark>student.name</mark>|student.familyname|student.age|
|---|---|---|---|
|1|Meimei|Han|12|
|2|Lily|Jaskson|15|

使用斜杠(/)作为分隔符：
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
## 空单元格

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

## 布尔值格式

布尔值格式指定了如何转换Excel中的布尔值：

1. `JSON true/false`：转换为JSON布尔值（`true`或`false`）
2. `字符串 "true"/"false"`：转换为字符串（`"true"`或`"false"`）
3. `数字 1/0`：转换为数字（`1`表示`TRUE`，`0`表示`FALSE`）

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

## 日期格式

日期格式指定了如何转换Excel中的日期值：

1. `从1900-01-01开始的天数`：转换为自1900-01-01以来的天数
2. `字符串, ISO 8601 (YYYY-MM-DDTHH:mm:ssZ)`：转换为ISO 8601格式的字符串

> 注意：日期格式功能仅在Excel数据表第一列添加$date$后缀时才有效。
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

## 更多功能

如果您已订阅并希望了解更多的功能，请发送邮件至he.yang@wtsolutions.cn

## 售后服务

如有任何问题或疑问，请通过邮件he.yang@wtsolutions.cn联系我们。我们将在24小时内尽快回复您，最迟不超过72小时。如果问题与您的订阅相关，请在邮件中包含您的`专业代码Pro Code`。
