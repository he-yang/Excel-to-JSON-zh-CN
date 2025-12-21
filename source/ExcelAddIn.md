# 3. Excel 加载项（在 Excel 中实现 Excel 转 JSON）

[English](https://excel-to-json.wtsolutions.cn/en/latest/ExcelAddIn.html)

WTSolutions 的 Excel 转 JSON 是一系列可将 Excel 转换为 JSON 的工具，支持转换扁平 JSON 和嵌套 JSON。它提供了“Excel 转 JSON”的全场景解决方案，包括 Excel 加载项、Web 应用程序、开放 API 和企业级 MCP 工具：

* [Web 应用：直接在网页浏览器中转换 Excel 为 JSON。](WebApp.md)
* <mark>Excel 加载项：在 Excel 中转换 Excel 为 JSON，与 Excel 环境无缝集成。</mark>（<-- 您当前所在位置）
* [WPS加载项：在WPS中转换Excel到JSON，与WPS环境无缝集成。](WPSAddIn.md)
* [API：通过 HTTPS POST 请求转换 Excel 为 JSON。](API.md)
* [MCP 服务：通过 AI 模型 MCP SSE/StreamableHTTP 请求转换 Excel 为 JSON。](MCP.md)

## 3.1 要求

* Excel 2013 Service Pack 1 或更高版本，
* Excel 2016 for Mac，
* Excel 2016 或更高版本，
* Excel Online，
* Office 365 等。

## 3.2 访问方式

* 在 Excel 2013/2016、Excel Online 或 Office 365 等中打开新的数据表
* **主页**选项卡或**插入**选项卡 > 加载项
* 在搜索框中，输入“Excel to JSON”
* 按照屏幕上的说明安装加载项，您将在**主页**选项卡上看到带有 Excel to JSON 图标的“Convert”按钮
* **主页**选项卡 > Excel to JSON > Convert
* 现在您已准备好使用此加载项

### 3.2.1 获取 Excel 加载项的视频指南（在 Excel 中侧载）

<iframe src="//player.bilibili.com/player.html?isOutside=true&aid=114329529684152&bvid=BV15ddBYLEaV&cid=29388573290&p=1" scrolling="no" border="0" frameborder="no" framespacing="0" allowfullscreen="true"></iframe>

[更多视频](https://s.wtsolutions.cn/images/videoqrcodes.png)



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

 <a name="Useadd-in"></a> 

## 3.3 使用方法

* 准备您的 Excel 表格
* 点击主页选项卡上带有 Excel to JSON 图标的“Convert”按钮
* 选择 Excel 数据源
    * 直接从 Excel 工作表中手动选择数据，或
    * 让 Excel to JSON 转换 Excel 中所有可见工作表，或 [专业版功能]
    * 让 Excel to JSON 转换 Excel 中所有工作表。[专业版功能]
* 设置[转换设置](profeatures.md)
* 点击“Go”按钮
* 之后您将在“Go”按钮下方看到转换后的 JSON。您可以通过以下几种方式将生成的 JSON 保存到本地计算机：
    * `Copy and Paste`（复制粘贴）。JSON 生成后，您将在文本区域中看到它们，您可以简单地将其复制粘贴到任何地方。
    * `Copy to Clipboard`（复制到剪贴板）。JSON 生成后，您可以找到“Copy to Clipboard”按钮，点击该按钮，JSON 将被复制到剪贴板。
    * `Save to File`（保存到文件）。（`Excel for Mac` 用户不可用）JSON 生成后，您可以找到“Save As”按钮，点击该按钮，系统将提示您将 JSON 保存到文件。

## 3.4 转换设置

详情请参考[转换设置](profeatures.md)。

## 3.5 Excel 加载项和 Web 应用的视频指南

<iframe src="//player.bilibili.com/player.html?isOutside=true&aid=114702621348629&bvid=BV16CN7zyE4d&cid=30558782233&p=1" scrolling="no" border="0" frameborder="no" framespacing="0" allowfullscreen="true"></iframe>

[更多视频](https://s.wtsolutions.cn/images/videoqrcodes.png)



