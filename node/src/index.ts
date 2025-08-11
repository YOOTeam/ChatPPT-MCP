#!/usr/bin/env node
import {Server} from "@modelcontextprotocol/sdk/server/index.js";
import {StdioServerTransport} from "@modelcontextprotocol/sdk/server/stdio.js";
import {
    CallToolRequestSchema,
    ListToolsRequestSchema,
    Tool,
} from "@modelcontextprotocol/sdk/types.js";
import fetch from "node-fetch";

function getApiKey(): string {
    const apiKey = process.env.API_KEY;
    if (!apiKey) {
        console.error("API_KEY environment variable is not set");
        process.exit(1);
    }
    return "Bearer " + apiKey;
}

const API_KEY = getApiKey();
const API_URL = "https://saas.api.yoo-ai.com"


interface ChatpptResponse {
    code: number;
    msg: string;
}

interface PptBuildResponse extends ChatpptResponse {
    data: {
        id: string
    }
}

interface PptQueryResponse extends ChatpptResponse {
    data: {
        images_url: object,
        id: string,
        introduce: string,
        ppt_title: string,
        status: number,
        process_url: string,
        preview_url: string,
        state_description: string
    }
}

interface PptDownloadResponse extends ChatpptResponse {
    data: {
        download_url: string,
    }
}

interface EditorPptResponse extends ChatpptResponse {
    data: {
        url: string,
    }
}

interface TaskResponse extends ChatpptResponse {
    data: {
        id: string
    }
}

interface DatabaseResponse extends ChatpptResponse {
    data: {
        data: string
    }
}


// Tool definitions
const BUILD_TOOL: Tool = {
    name: "build_ppt",
    description: "根据描述的文本或markdown，执行生成任务。当返回PPT-ID时，表示生成任务成功，可以调用query_ppt工具查询生成进度",
    inputSchema: {
        type: "object",
        properties: {
            text: {
                type: "string",
                description: "输入描述的文本或markdown，生成PPT"
            }
        },
        required: ["text"]
    }
};

const QUERY_TOOL: Tool = {
    name: "query_ppt",
    description: "根据PPT任务ID查询异步生成结果，status=1表示还在生成中，应该继续轮训该查询，status=2表示成功，status=3表示失败；process_url表示预览的url地址，不断轮训请求直至成功或失败;\n" +
        "        当成功后使用默认浏览器打开ppt地址并下载PPT和生成编辑器地址；",
    inputSchema: {
        type: "object",
        properties: {
            ppt_id: {
                type: "string",
                description: "PPT-ID"
            }
        },
        required: ["ppt_id"]
    }
};

const DOWNLOAD_PPT_TOOL: Tool = {
    name: "download_ppt",
    description: "根据PPT任务ID生成 PPT 下载地址",
    inputSchema: {
        type: "object",
        properties: {
            id: {
                type: "string",
                description: "PPT-ID"
            }
        },
        required: ["id"]
    }
};

const EDITOR_PPT_TOOL: Tool = {
    name: "editor_ppt",
    description: "根据PPT任务ID生成PPT编辑器界面URL",
    inputSchema: {
        type: "object",
        properties: {
            id: {
                type: "string",
                description: "PPT-ID"
            }
        },
        required: ["id"]
    }
};

const REPLACE_PPT_TOOL: Tool = {
    name: "replace_ppt",
    description: "根据任务PPT-ID执行随机替换PPT模板，并返回新的任务PPT-ID。",
    inputSchema: {
        type: "object",
        properties: {
            ppt_id: {
                type: "string",
                description: "PPT-ID"
            }
        },
        required: ["id"]
    }
}

const SET_FONT_PPT_TOOL: Tool = {
    name: "set_font_ppt",
    description: "根据PPT-ID执行设置PPT字体，参照给定的字体名称，如黑体、宋体、仿宋、幼圆、楷体、隶书等进行设置，，并返回新的PPT-ID。",
    inputSchema: {
        type: "object",
        properties: {
            ppt_id: {
                type: "string",
                description: "PPT-ID"
            },
            font_name: {
                type: "string",
                description: "字体名，如黑体、宋体、仿宋、幼圆、楷体、隶书"
            }
        },
        required: ["id"]
    }
}

const SET_ANIM_PPT_TOOL: Tool = {
    name: "set_anim_ppt",
    description: "根据PPT-ID执行设置动画,可以支持参照给出的任务PPT-ID设置或者取消用户PPT的动画效果。",
    inputSchema: {
        type: "object",
        properties: {
            ppt_id: {
                type: "string",
                description: "PPT-ID"
            },
            set_anim: {
                type: "string",
                default: "1",
                description: "设置动画还是取消动画"
            }
        },
        required: ["id"]
    }
}

const CREATE_NOTE_PPT_TOOL: Tool = {
    name: "create_note_ppt",
    description: "参照给出的任务PPT-ID自动为用户的ppt生成全文演讲稿，并会返回新的PPT-ID。",
    inputSchema: {
        type: "object",
        properties: {
            ppt_id: {
                type: "string",
                description: "PPT-ID"
            }
        },
        required: ["id"]
    }
}

const ADD_SLIDES_PPT_TOOL: Tool = {
    name: "add_slides_ppt",
    description: "参照给出的任务PPT-ID，给对应的文档新增或插入新的PPT页面，可以指定对应的页面类型，返回新的PPT-ID。",
    inputSchema: {
        type: "object",
        properties: {
            ppt_id: {
                type: "string",
                description: "PPT-ID",
                required: true
            },
            slide_text: {
                type: "string",
                description: "slide_text，用户指定插入页数",
                required: true
            },
            slide_type: {
                type: "string",
                description: "slide_type，指定生成页面类型，可以是“封面页，目录页，章节页，内容页，致谢页”，默认为内容页。",
                default: "内容页"
            }
        },
        required: ["id"]
    }
}

const CREATE_OUTLINE_PPT_TOOL: Tool = {
    name: "create_outline_ppt",
    description: "根据用户输入的内容ppt_text，实时生成大纲内容，直接返回大纲文本内容。",
    inputSchema: {
        type: "object",
        properties: {
            ppt_text: {
                type: "string",
                description: "根据用户输入的内容ppt_text，实时生成大纲内容，直接返回大纲文本内容。",
                required: true
            }
        },
        required: ["ppt_text"]
    }
}

const CREATE_TEMPLATE_COVER_TOOL: Tool = {
    name: "create_template_cover_ppt",
    description: "根据用户输入的内容ppt_text，实时生成与渲染对应的模板，返回对应的模板ID和对应的渲染图片，支持用户指定颜色ppt_color和风格ppt_style，默认随机；也可以指定返回的数量，默认为4个。\n",
    inputSchema: {
        type: "object",
        properties: {
            ppt_text: {
                type: "string",
                description: "根据用户输入的内容ppt_text，实时生成大纲内容，直接返回大纲文本内容。",
            },
            ppt_color: {
                type: "string",
                description: "Template-color，指定生成的模板风格，可以为空，表示随机；也可以从\"科技风\",\"商务风\",\"小清新\",\"极简风\",\"中国风\",\"可爱卡通\"进行执行，可选参数。"
            },
            ppt_style: {
                type: "string",
                description: "Template-style,指定生成模板的颜色，可以为空，表示随机；也可以从\"紫色\",\"红色\",\"橙色\",\"黄色\",\"绿色\",\"青色\",\"蓝色\",\"粉色\",\"灰色\"进行指定，可选参数。"
            },
            ppt_num: {
                type: "int",
                description: "Template-num，指定生成模板数量，默认为4",
                default: 4
            }
        },
        required: ["ppt_text"]
    }
}

const BUILD_PPT_BY_FILE: Tool = {
    name: "build_ppt_by_file",
    description: "根据用户上传的文件（给出文件url地址），执行生成PPT的任务。当返回PPT-ID时，表示生成任务成功，可以调用query_ppt工具查询生成进度和预览URL",
    inputSchema: {
        type: "object",
        properties: {
            file_url: {
                type: "string",
                description: "用户给定的文件url地址，可以支持包括MarkDown、word、PDF、XMind、FreeMind、TXT 等文档文件",
            }
        },
        required: ["file_url"]
    }
}

const REPLACE_TEMPLATE_COVER_TOOL: Tool = {
    name: "replace_template_cover_ppt",
    description: "根据任务PPT-ID执行替换为用户指定（根据cover_id）的模板，并返回新的任务PPT-ID。",
    inputSchema: {
        type: "object",
        properties: {
            ppt_id: {
                type: "string",
                description: "PPT-ID",
            },
            cover_id: {
                type: "string",
                description: "cover_id，用户指定的模板id，需要通过ppt_create_template_coverImage 进行生成。"
            }
        },
        required: ["ppt_id", "cover_id"]
    }
}

// 生成ppt
async function handlePptBuild(text: string) {
    const url = new URL(API_URL + "/mcp/ppt/ppt-create");
    let params = JSON.stringify({
        "text": text,
    })
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": API_KEY
        },
        body: params,
    });
    const data = await response.json() as PptBuildResponse;
    if (data.code !== 200) {
        return {
            content: [{
                type: "text",
                text: `BuildPpt failed: ${data.code} : ${data.msg}`,
            }],
            isError: true
        };
    }

    return {
        content: [{
            type: "text",
            text: JSON.stringify({
                id: data.data.id,
            }, null, 2)
        }],
        isError: false
    };
}


// 查询ppt
async function handleQuery(ppt_id: string) {
    const url = new URL(API_URL + "/mcp/ppt/ppt-result");
    url.searchParams.append("id", ppt_id);

    const response = await fetch(url.toString(), {
        method: "GET",
        headers: {
            "Content-Type": "application/json",
            "Authorization": API_KEY
        }
    });
    const data = await response.json() as PptQueryResponse;
    if (data.code !== 200) {
        return {
            content: [{
                type: "text",
                text: `Ppt failed: ${data.code} : ${data.msg}`,
            }],
            isError: true
        };
    }

    return {
        content: [{
            type: "text",
            text: JSON.stringify({
                id: data.data.id,
                images_url: data.data.images_url,
                introduce: data.data.introduce,
                ppt_title: data.data.ppt_title,
                status: data.data.status,
                process_url: data.data.process_url,
                preview_url: data.data.preview_url,
                state_description: data.data.state_description,
            }, null, 2)
        }],
        isError: false
    };
}


// 下载ppt
async function handleDownloadPpt(id: string) {
    const url = new URL(API_URL + "/mcp/ppt/ppt-download");
    url.searchParams.append("id", id);

    const response = await fetch(url.toString(), {
        method: "GET",
        headers: {
            "Content-Type": "application/json",
            "Authorization": API_KEY
        }
    });
    const data = await response.json() as PptDownloadResponse;
    if (data.code !== 200) {
        return {
            content: [{
                type: "text",
                text: `Ppt failed: ${data.code} : ${data.msg}`,
            }],
            isError: true
        };
    }

    return {
        content: [{
            type: "text",
            text: JSON.stringify({
                download_url: data.data.download_url,
            }, null, 2)
        }],
        isError: false
    };
}


// 编辑器ppt
async function handleEditorPpt(id: string) {
    const url = new URL(API_URL + "/mcp/ppt/ppt-editor");
    let params = JSON.stringify({
        "id": id,
    })

    const response = await fetch(url.toString(), {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": API_KEY
        },
        body: params,
    });
    const data = await response.json() as EditorPptResponse;
    if (data.code !== 200) {
        return {
            content: [{
                type: "text",
                text: `Ppt failed: ${data.code} : ${data.msg}`,
            }],
            isError: true
        };
    }

    return {
        content: [{
            type: "text",
            text: JSON.stringify({
                url: data.data.url,
            }, null, 2)
        }],
        isError: false
    };
}


// 替换/更换模板
async function handleReplacePpt(id: string) {
    const url = new URL(API_URL + "/mcp/ppt/ppt-create-task");
    let params = JSON.stringify({
        "id": id,
    })
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": API_KEY
        },
        body: params,
    });
    const data = await response.json() as TaskResponse;
    if (data.code !== 200) {
        return {
            content: [{
                type: "text",
                text: `BuildPpt failed: ${data.code} : ${data.msg}`,
            }],
            isError: true
        };
    }

    return {
        content: [{
            type: "text",
            text: JSON.stringify({
                id: data.data.id,
            }, null, 2)
        }],
        isError: false
    };
}


// 替换字体
async function handleSetFontPpt(id: string, font_name: string) {
    const url = new URL(API_URL + "/mcp/ppt/ppt-create-task");
    let params = JSON.stringify({
        "id": id,
        "font_name": font_name
    })
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": API_KEY
        },
        body: params,
    });
    const data = await response.json() as TaskResponse;
    if (data.code !== 200) {
        return {
            content: [{
                type: "text",
                text: `BuildPpt failed: ${data.code} : ${data.msg}`,
            }],
            isError: true
        };
    }

    return {
        content: [{
            type: "text",
            text: JSON.stringify({
                id: data.data.id,
            }, null, 2)
        }],
        isError: false
    };
}


// 生成动画
async function handleSetAnimPpt(id: string, set_anim: string) {
    const url = new URL(API_URL + "/mcp/ppt/ppt-create-task");
    let params = JSON.stringify({
        "id": id,
        "set_anim": set_anim
    })
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": API_KEY
        },
        body: params,
    });
    const data = await response.json() as TaskResponse;
    if (data.code !== 200) {
        return {
            content: [{
                type: "text",
                text: `BuildPpt failed: ${data.code} : ${data.msg}`,
            }],
            isError: true
        };
    }

    return {
        content: [{
            type: "text",
            text: JSON.stringify({
                id: data.data.id,
            }, null, 2)
        }],
        isError: false
    };
}


// 生成演讲稿
async function handleCreateNotePpt(id: string) {
    const url = new URL(API_URL + "/mcp/ppt/ppt-create-task");
    let params = JSON.stringify({
        "id": id,
        "note": "1"
    })
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": API_KEY
        },
        body: params,
    });
    const data = await response.json() as TaskResponse;
    if (data.code !== 200) {
        return {
            content: [{
                type: "text",
                text: `BuildPpt failed: ${data.code} : ${data.msg}`,
            }],
            isError: true
        };
    }

    return {
        content: [{
            type: "text",
            text: JSON.stringify({
                id: data.data.id,
            }, null, 2)
        }],
        isError: false
    };
}

// 新增页面
async function handleAddSlidesPPT(id: string, slide_text: string, slide_type: string) {
    const url = new URL(API_URL + "/mcp/ppt/ppt-page");
    let params = JSON.stringify({
        "id": id,
        "slide_text": slide_text,
        "slide_type": slide_type
    })
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": API_KEY
        },
        body: params,
    });
    const data = await response.json() as TaskResponse;
    if (data.code !== 200) {
        return {
            content: [{
                type: "text",
                text: `BuildPpt failed: ${data.code} : ${data.msg}`,
            }],
            isError: true
        };
    }

    return {
        content: [{
            type: "text",
            text: JSON.stringify({
                id: data.data.id,
            }, null, 2)
        }],
        isError: false
    };
}


// 生成大纲
async function handleCreateOutLinePPT(ppt_text: string) {
    const url = new URL(API_URL + "/mcp/ppt/ppt-structure");
    let params = JSON.stringify({
        "text": ppt_text,
    })
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": API_KEY
        },
        body: params,
    });
    const data = await response.json() as DatabaseResponse;
    if (data.code !== 200) {
        return {
            content: [{
                type: "text",
                text: `BuildPpt failed: ${data.code} : ${data.msg}`,
            }],
            isError: true
        };
    }

    return {
        content: [{
            type: "text",
            text: JSON.stringify({
                id: data.data,
            }, null, 2)
        }],
        isError: false
    };
}


// 生成封面图
async function handleCreateTemplateCover(ppt_text: string, ppt_color: string, ppt_style: string, ppt_num: number) {
    const url = new URL(API_URL + "/mcp/ppt/ppt-cover");
    let params = JSON.stringify({
        "title": ppt_text,
        "count": ppt_num,
        "color": ppt_color,
        "style": ppt_style
    })
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": API_KEY
        },
        body: params,
    });
    const data = await response.json() as DatabaseResponse;
    if (data.code !== 200) {
        return {
            content: [{
                type: "text",
                text: `BuildPpt failed: ${data.code} : ${data.msg}`,
            }],
            isError: true
        };
    }

    return {
        content: [{
            type: "text",
            text: JSON.stringify({
                id: data.data,
            }, null, 2)
        }],
        isError: false
    };
}


// 替换指定模板
async function handleReplaceTemplateCover(ppt_id: string, cover_id: string) {
    const url = new URL(API_URL + "/mcp/ppt/ppt-create-task");
    let params = JSON.stringify({
        "ppt_id": ppt_id,
        "cover_id": cover_id
    })
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": API_KEY
        },
        body: params,
    });
    const data = await response.json() as DatabaseResponse;
    if (data.code !== 200) {
        return {
            content: [{
                type: "text",
                text: `BuildPpt failed: ${data.code} : ${data.msg}`,
            }],
            isError: true
        };
    }

    return {
        content: [{
            type: "text",
            text: JSON.stringify({
                id: data.data,
            }, null, 2)
        }],
        isError: false
    };
}

// 通过文件url生成PPT
async function handleBuildPptByFile(file_url: string) {
    const url = new URL(API_URL + "/mcp/ppt/ppt-create-file");
    let params = JSON.stringify({
        "file_url": file_url
    })
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": API_KEY
        },
        body: params,
    });
    const data = await response.json() as DatabaseResponse;
    if (data.code !== 200) {
        return {
            content: [{
                type: "text",
                text: `BuildPpt failed: ${data.code} : ${data.msg}`,
            }],
            isError: true
        };
    }

    return {
        content: [{
            type: "text",
            text: JSON.stringify({
                id: data.data,
            }, null, 2)
        }],
        isError: false
    };
}


// Create an MCP server
const server = new Server(
    {
        name: "mcp-server/chatppt",
        version: "1.0.0",
    },
    {
        capabilities: {
            tools: {},
        },
    },
);

const MAPS_TOOLS = [
    BUILD_TOOL,
    QUERY_TOOL,
    DOWNLOAD_PPT_TOOL,
    EDITOR_PPT_TOOL,
    REPLACE_PPT_TOOL,
    SET_FONT_PPT_TOOL,
    SET_ANIM_PPT_TOOL,
    CREATE_NOTE_PPT_TOOL,
    ADD_SLIDES_PPT_TOOL,
    CREATE_OUTLINE_PPT_TOOL,
    CREATE_TEMPLATE_COVER_TOOL,
    REPLACE_TEMPLATE_COVER_TOOL,
    BUILD_PPT_BY_FILE
] as const;

// Set up request handlers
server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: MAPS_TOOLS,
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
    try {
        switch (request.params.name) {
            case "build_ppt": {
                const {text} = request.params.arguments as { text: string };
                return await handlePptBuild(text);
            }

            case "query_ppt": {
                const {ppt_id} = request.params.arguments as { ppt_id: string };
                return await handleQuery(ppt_id);
            }

            case "download_ppt": {
                const {id} = request.params.arguments as { id: string };
                return await handleDownloadPpt(id);
            }

            case "editor_ppt": {
                const {id} = request.params.arguments as { id: string };
                return await handleEditorPpt(id);
            }

            case "replace_ppt": {
                const {id} = request.params.arguments as { id: string };
                return await handleReplacePpt(id);
            }

            case "set_font_ppt": {
                const {id, font_name} = request.params.arguments as { id: string, font_name: string };
                return await handleSetFontPpt(id, font_name)
            }

            case "set_anim_ppt": {
                const {id, set_anim} = request.params.arguments as { id: string, set_anim: string };
                return await handleSetAnimPpt(id, set_anim)
            }

            case "create_note_ppt": {
                const {id} = request.params.arguments as { id: string };
                return await handleCreateNotePpt(id)
            }

            case "add_slides_ppt": {
                const {id, slide_text, slide_type} = request.params.arguments as {
                    id: string,
                    slide_type: string,
                    slide_text: string
                };
                return await handleAddSlidesPPT(id, slide_text, slide_type)
            }

            case "create_outline_ppt": {
                const {ppt_text} = request.params.arguments as { ppt_text: string };
                return await handleCreateOutLinePPT(ppt_text)
            }

            case "create_template_cover_ppt": {
                const {ppt_text, ppt_num, ppt_style, ppt_color} = request.params.arguments as {
                    ppt_text: string,
                    ppt_num: number,
                    ppt_style: string,
                    ppt_color: string
                };
                return await handleCreateTemplateCover(ppt_text, ppt_color, ppt_style, ppt_num)
            }

            case "replace_template_cover_ppt": {
                const {id, cover_id} = request.params.arguments as { id: string, cover_id: string };
                return await handleReplaceTemplateCover(id, cover_id)
            }

            case "build_ppt_by_file": {
                const {file_url} = request.params.arguments as { file_url: string };
                return await handleBuildPptByFile(file_url)
            }

            default:
                return {
                    content: [{
                        type: "text",
                        text: `Unknown tool: ${request.params.name}`
                    }],
                    isError: true
                };
        }
    } catch (error) {
        return {
            content: [{
                type: "text",
                text: `Error: ${error instanceof Error ? error.message : String(error)}`
            }],
            isError: true
        };
    }
});


async function runServer() {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error("Chatppt MCP Server running on stdio");
}

runServer().catch((error) => {
    console.error("Fatal error running server:", error);
    process.exit(1);
});