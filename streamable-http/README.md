# 一、使用说明

为了进一步提升ChatPPT的MCP Serverr体验，除了你可以选择使用标准的STDIO方式进行下载源码进行部署和使用外，我们也将为所有用户开放最新的MCP “Streamable HTTP”方式，即只需要在可支持的MCP客户端或者工具上，配置对应的ChatPPT Server服务链接，即可轻松体验，无需任何代码环境。

---
# 二、体验流程
## 1.安装MCP 客户端
建议使用最新版本的 Cursor 客户端（最新版，最低版本为v0.5），下载地址 https://cursor.com/cn。
为方便理解，本示例文档后文将以Cursor作为示例工具。

 ## 2.配置ChatPPT Server
 进入 Cursor 设置界面，并配置ChatPPT MCP Server，注意：一定要选择 Streamable HTTP  方式。
 
 配置实例参考如下：

 ```json
{
    "chatppt-http": {
        "autoApprove": [
            "query_ppt",
            "build_ppt",
            "ppt_replace_template",
            "ppt_set_color",
            "download_ppt",
            "editor_ppt"
            ],
        "disabled": false,
        "timeout": 60,
        "url": "http://mcp.test.yoo-ai.com/mcp?key=YOO_PPT_KEY",
        "transportType": "streamable-http"
    }
}
```
![添加server](https://image.yoojober.com/users/2025-08/68ad11463b946.png)

>**说明**：其中的“YOO_PPT_KEY” 需要替换你对应的ChatPPT 访问API_Key，可以访问 [ChatPPT服务指南](https://wiki.yoo-ai.com/mcp/McpServe/serve1.3.html)了解对应内容。

## 3.查看配置Server状态
配置结束后，返回Cursor查看对应的server配置状态，正常显示绿色状态，且可以看到对应Server对应的Tools。
![查看server状态](https://image.yoojober.com/users/2025-08/68ad11af310c5.png)

## 4.配置客户端的大模型
选择配置 Cursor 大模型让你拥有更好的服务体验，建议选择 claude-4-sonnet等模型。
![配置大模型](https://image.yoojober.com/users/2025-08/68ad1169652fc.png)

## 5.使用与测试生成PPT
（1）按下 CTRL/CMD + L 快捷键，即可在编辑器右侧打开对话框

（2）选择在Cursor的交互模式选择，选择 Agent 方式

（3）输入你要测试生成的内容，例如：帮我生成一份关于智能体技术发展的PPT

（4）后续参照AI对话进行使用即可，你将获得实时预览界面、在线编辑页面，以及pptx文件的下载地址。

> **说明**：如果在使用过程中，AI会有部分针对一些MCP的Tool调用过程的确认，可以根据实际情况选择“Run Tool” 即可。
![过程1](https://image.yoojober.com/users/2025-08/68ad134b11936.png)

![过程2](https://image.yoojober.com/users/2025-08/68ad1376b8c54.png)

**【实时预览效果】**

![结果1](https://image.yoojober.com/users/2025-08/68ad150e7d7f8.png)

**【在线编辑效果】**

![结果2](https://image.yoojober.com/users/2025-08/68ad1519d332d.png)


**【下载pptx文件，兼容Office】**

![结果3](https://image.yoojober.com/users/2025-08/68ad15235da0d.png)


---

# 三、Tools功能清单
为方便你的使用，目前ChatPPT MCP Server已经支持的Tool有：

>check_api_key、query_ppt、build_ppt、build_ppt_by_file、build_thesis_ppt、ppt_replace_template、ppt_set_color、text_build_ppt、download_ppt、editor_ppt、ppt_set_font_name、ppt_set_anim、ppt_create_note、ppt_add_slides、ppt_create_outline_text、ppt_create_template_cover_image、ppt_replace_user_select_template

ChatPP对应的服务Tool未来将持续更新和开放更多，部分核心Tool介绍如下。

- ## “build_ppt”（生成PPT）
    根据输入的主题或文本内容进行PPT。
- ## “query_pptt”（查询状态）
    实时查询生成状态，如果有生成的页面，将会实时返回预览幻灯片图。
- ## “download_ppt”（获取下载PPT地址）
    获取下载的pptx文件地址url。
- ## “editor_ppt”（生成PPT）
    返回对应内容的在线编辑地址。
- ## “build_ppt_by_file”（文件生成PPT）
    根据指定的文件url生成对应的内容。
- ## “ppt_replace_template”（跟换PPT模板）
    随机更换当前PPT模板风格。
- ## “ppt_create_note”（生成演讲稿）
    为当前文档生成演讲稿。
- ## “ppt_set_color”、“ppt_set_font_name”、“ppt_set_anim”（设置颜色、字体、动画）
    为当前文档重新设置主题色、字体和动画等
 - ## “ppt_replace_template”（文件生成PPT）
    根据指定的文件url生成对应的内容。

 更多介绍可以参考[完整Tool文档](https://wiki.yoo-ai.com/mcp/McpServe/serve1.2.html)