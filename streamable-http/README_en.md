# User Guide

To enhance your ChatPPT MCP Server experience, you can either deploy using the standard STDIO method or leverage the new **"Streamable HTTP"** method. This approach allows configuration directly in compatible MCP clients/tools without requiring any code environment setup.

---
## 1. Install MCP Client
We recommend using the latest version of **Cursor client** (minimum v0.5):  
📥 Download: https://cursor.com/cn  
*Note: This guide uses Cursor for demonstration purposes.*

## 2. Configure ChatPPT Server
1. Navigate to Cursor's settings interface
2. Configure ChatPPT MCP Server with **"Streamable HTTP"** selected
3. Use the following configuration template:

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
        "url": "http://mcp.yoo-ai.com/mcp?key=YOO_PPT_KEY",
        "transportType": "streamable-http"
    }
}
```


> **Important**: Replace `YOO_PPT_KEY` with your actual ChatPPT API Key.  
> 🔗 Learn more: [Guide Book](https://wiki.yoo-ai.com/mcp/McpServe/serve1.3.html)

## 3. Check Server Status
After configuration, ensure the server status shows **green** in Cursor and displays available tools:
![Server Status](https://image.yoojober.com/users/2025-08/68ad11af310c5.png)


## 4. Configure AI Model
For optimal performance, select powerful models like **claude-4-sonnet**:
![Confignation](https://image.yoojober.com/users/2025-08/68ad1169652fc.png)


## 5. Generate and Test PPTs
1. Press **CTRL/CMD + L** to open chat panel
2. Select **Agent mode**
3. Enter your request (e.g., _"Generate a PPT about intelligent agent technology development"_)
4. Follow AI instructions and confirm tool executions when prompted
![过程1](https://image.yoojober.com/users/2025-08/68ad134b11936.png)

![过程2](https://image.yoojober.com/users/2025-08/68ad1376b8c54.png)



**Real-time Preview**  

![结果1](https://image.yoojober.com/users/2025-08/68ad150e7d7f8.png)

**Online Editor**  

![结果2](https://image.yoojober.com/users/2025-08/68ad1519d332d.png)


**Download PPTX (Office Compatible)**  

![结果3](https://image.yoojober.com/users/2025-08/68ad15235da0d.png)

---
# Supported Tools
ChatPPT MCP Server currently supports these tools:

> `check_api_key`, `query_ppt`, `build_ppt`, `build_ppt_by_file`, `build_thesis_ppt`, `ppt_replace_template`, `ppt_set_color`, `text_build_ppt`, `download_ppt`, `editor_ppt`, `ppt_set_font_name`, `ppt_set_anim`, `ppt_create_note`, `ppt_add_slides`, `ppt_create_outline_text`, `ppt_create_template_cover_image`, `ppt_replace_user_select_template`

## Core Tools Overview
- **`build_ppt`**: Generates PPT from input topics/text
- **`query_ppt`**: Checks generation status and returns slide previews
- **`download_ppt`**: Provides PPTX download URL
- **`editor_ppt`**: Returns online editing URL
- **`build_ppt_by_file`**: Creates PPT from specified file URL
- **`ppt_replace_template`**: Randomizes PPT template style
- **`ppt_create_note`**: Generates speech notes for current document
- **`ppt_set_color/font_name/anim`**: Modifies theme colors, fonts, and animations

📚 Full documentation: [完整Tool文档](https://wiki.yoo-ai.com/mcp/McpServe/serve1.2.html)
