# coding: utf-8
"""
Chatppt MCP Server
"""
import os
from typing import List

import httpx
from mcp.server.fastmcp import FastMCP
from pydantic import Field

# 创建MCP服务器实例
mcp = FastMCP("Chatppt Server", log_level="ERROR")
# Chatppt API Base URL
API_BASE = "https://saas.api.yoo-ai.com"
# 用户API Key
API_PPT_KEY = os.getenv('API_PPT_KEY')


# 查询apikey
def check_api_key():
    """检查 API_PPT_KEY 是否已设置"""
    if not API_PPT_KEY:
        raise ValueError("API_PPT_KEY 环境变量未设置")
    return API_PPT_KEY


@mcp.tool()
async def check():
    """查询用户当前配置token"""
    return os.getenv('API_PPT_KEY')


# 注册工具的装饰器，可以很方便的把一个函数注册为工具
@mcp.tool()
async def query_ppt(ppt_id: str = Field(description="PPT-ID")) -> str:
    """
    Name:
        查询PPT生成进度
    Description:
        根据PPT任务ID查询异步生成结果，status=1表示还在生成中，应该继续轮训该查询，status=2表示成功，status=3表示失败；process_url表示预览的url地址，不断轮训请求直至成功或失败;
        当成功后使用默认浏览器打开ppt地址并调用download_ppt工具下载PPT和工具editor_ppt生成编辑器地址；
    Args:
        ppt_id: PPT-ID
    Returns:
        PPT信息的描述
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-result'

        async with httpx.AsyncClient() as client:
            response = await client.get(
                url,
                params={'id': ppt_id},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=30
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except KeyError as e:
        raise Exception(f"Failed to parse response: {str(e)}") from e
    except httpx.HTTPError as e:
        raise Exception(f"HTTP request failed: {str(e)}") from e


# 主题生成PPT
@mcp.tool()
async def build_ppt(
        theme: str = Field(description="描述生成主题"),
) -> str:
    """
    Name:
        PPT生成。当用户需要生成PPT时，调用此工具
    Description:
        根据主题生成ppt。当返回PPT-ID时，表示生成任务成功，可以调用query_ppt工具查询生成进度和预览URL
    Args:
        theme: 输入描述的描述生成主题或markdown，生成PPT
    Returns:
        PPT-ID
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-create'

        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                data={'text': theme},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=30
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP请求失败: {str(e)}") from e
    except ValueError as e:
        raise Exception(str(e)) from e
    except Exception as e:
        raise Exception(f"PPT生成失败: {str(e)}") from e


@mcp.tool()
async def text_build_ppt(
        text: str = Field(description="根据长文本（50字以上）"),
) -> str:
    """
    Name:
        根据长文本（50字以上）生成PPT。
    Description:
        根据长文本（50字以上）生成PPT。当返回PPT-ID时，表示生成任务成功，可以调用query_ppt工具查询生成进度和预览URL
    Args:
        text: 输入描述的文本（50字以上）或markdown，生成PPT
    Returns:
        PPT-ID
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-create'

        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                data={'text': text},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=30
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP请求失败: {str(e)}") from e
    except ValueError as e:
        raise Exception(str(e)) from e
    except Exception as e:
        raise Exception(f"PPT生成失败: {str(e)}") from e


# 文件生成PPT
@mcp.tool()
async def build_ppt_by_file(
        file_url: str = Field(description="文件地址"),
) -> str:
    """
    Name:
        文件生成PPT。
    Description:
        根据用户上传的文件（给出文件url地址），执行生成PPT的任务。当返回PPT-ID时，表示生成任务成功，可以调用query_ppt工具查询生成进度和预览URL
    Args:
        file_url: 用户给定的文件url地址，可以支持包括MarkDown、word、PDF、XMind、FreeMind、TXT 等文档文件
    Returns:
        PPT-ID
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-create-file'

        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                data={'file_url': file_url},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=30
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP请求失败: {str(e)}") from e
    except ValueError as e:
        raise Exception(str(e)) from e
    except Exception as e:
        raise Exception(f"PPT生成失败: {str(e)}") from e


# 论文生成PPT
@mcp.tool()
async def build_thesis_ppt(
        file_url: str = Field(description="文件地址"),
) -> str:
    """
    Name:
        论文文件生成答辩PPT。
    Description:
        根据用户上传的文件（给出文件url地址，），执行生成PPT的任务。当返回PPT-ID时，表示生成任务成功，可以调用query_ppt工具查询生成进度和预览URL
    Args:
        file_url: 用户的论文文件地址url，仅支持pdf、word、pdf三种文件
    Returns:
        PPT-ID
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-create-thesis'

        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                data={'file_key': file_url},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=30
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP请求失败: {str(e)}") from e
    except ValueError as e:
        raise Exception(str(e)) from e
    except Exception as e:
        raise Exception(f"PPT生成失败: {str(e)}") from e


# 下载PPT文件
@mcp.tool()
async def download_ppt(
        ppt_id: str = Field(description="PPT-ID")
) -> str:
    """
    Name:
        当PPT生成完成后，生成下载PPT的地址，方便用户下载到本地。
    Description:
        获取完整生成PPT文件的下载地址，仅当PPT生成完成后，才生成此下载地址。
    Args:
        ppt_id: PPT-ID
    Returns:
        PPT下载地址URL
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-download'

        async with httpx.AsyncClient() as client:
            response = await client.get(
                url,
                params={'id': ppt_id},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=60
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP request failed: {str(e)}") from e


# 在线编辑PPT
@mcp.tool()
async def editor_ppt(
        ppt_id: str = Field(description="PPT-ID")
) -> str:
    """
    Name:
         基于生成后的文件，打开并展示pptx文件，方便进行在线编辑与浏览查看。
    Description:
        通过PPT-ID生成PPT编辑器界面URL
    Args:
        ppt_id: PPT-ID
    Returns:
        PPT编辑器地址URL
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-editor'

        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                data={'id': ppt_id},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=60
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP request failed: {str(e)}") from e


# 替换/更换模板
@mcp.tool()
async def ppt_replace_template(ppt_id: str = Field(description="PPT-ID")) -> str:
    """
    Name:
        更改替换为PPT模板。
    Description:
        根据任务PPT-ID执行随机替换PPT模板，并返回新的任务PPT-ID，可以调用query_ppt工具查询生成进度和预览URL
    Args:
        ppt_id: PPT-ID
    Returns:
        新的PPT-ID
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-create-task'

        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                data={'id': ppt_id},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=60
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP request failed: {str(e)}") from e


# 设置ppt主题色
@mcp.tool()
async def ppt_set_color(
        ppt_id: str = Field(description="PPT-ID"),
        color: str = Field(description="PPT-color")
) -> str:
    """
    Name:
        更改设置PPT主题色。
    Description:
        根据PPT-ID执行设置更换主题色，可以参照用给定的颜色空间名、或者颜色值等，进行设置PPT主题色，并返回新的PPT-ID，可以调用query_ppt工具查询生成进度和预览URL
    Args:
        ppt_id(str): PPT-ID
        color(str)：PPT-color,设置的主题色，可以支持颜色空间名称如"紫色","红色","橙色","黄色","绿色","青色","蓝色","粉色",也可以支持#xxxxxh或者RGB颜色值。
    Returns:
        新的PPT-ID
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-create-task'

        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                data={'id': ppt_id, "color": color},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=60
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP request failed: {str(e)}") from e


# 设置ppt字体
@mcp.tool()
async def ppt_set_font_name(
        ppt_id: str = Field(description="PPT-ID"),
        font_name: str = Field(description="font_name")
) -> str:
    """
    Name:
        设置更改字体。
    Description:
        根据PPT-ID执行设置PPT字体，参照给定的字体名称，如黑体、宋体、仿宋、幼圆、楷体、隶书等进行设置，，并返回新的PPT-ID，可以调用query_ppt工具查询生成进度和预览URL
    Args:
        ppt_id(str): PPT-ID
        font_name(str)：字体名，如黑体、宋体、仿宋、幼圆、楷体、隶书
    Returns:
        新的PPT-ID
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-create-task'

        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                data={'id': ppt_id, "font_name": font_name},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=60
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP request failed: {str(e)}") from e


# 设置ppt动画
@mcp.tool()
async def ppt_set_anim(ppt_id: str = Field(description="PPT-ID"),
                       set_anim: str = Field(default="1", description="设置动画还是取消动画")) -> str:
    """
    Name:
        更改设置动画。
    Description:
        根据PPT-ID执行设置动画,可以支持参照给出的任务PPT-ID设置或者取消用户PPT的动画效果。
    Args:
        ppt_id: PPT-ID
        set_anim(str)：是否设置，默认为“1”表示设置动画，“0”标识取消动画。
    Returns:
        新的PPT-ID
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-create-task'

        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                data={'id': ppt_id, "transition": set_anim},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=60
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP request failed: {str(e)}") from e


# 生成ppt演讲稿
@mcp.tool()
async def ppt_create_note(ppt_id: str = Field(description="PPT-ID")) -> str:
    """
    Name:
        生成演讲稿。
    Description:
        参照给出的任务PPT-ID自动为用户的ppt生成全文演讲稿，并会返回新的PPT-ID，可以调用query_ppt工具查询生成进度和预览URL
    Args:
        ppt_id: PPT-ID
    Returns:
        新的PPT-ID
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-create-task'

        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                data={'id': ppt_id, "note": "1"},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=60
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP request failed: {str(e)}") from e


# 插入新增ppt页面
@mcp.tool()
async def ppt_add_slides(ppt_id: str = Field(description="PPT-ID"),
                         slide_text: str = Field(description="slide-themeText"),
                         slide_type: str = Field(default="内容页", description="slide_type"), ) -> str:
    """
    Name:
        给生成后的PPT，插入或者新增单页幻灯片。
    Description:
        参照给出的任务PPT-ID，给对应的文档新增或插入新的PPT页面，可以指定对应的页面类型，返回新的PPT-ID，可以调用query_ppt工具查询生成进度和预览URL
    Args:
        ppt_id: PPT-ID，必须参数。
        slide_text(str): slide_text，用户指定插入页数，必须参数。
        slide_type(str): slide_type，指定生成页面类型，可以是“封面页，目录页，章节页，内容页，致谢页”，默认为内容页。
    Returns:
        新的PPT-ID
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-page'

        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                data={'id': ppt_id, 'slide_text': slide_text, 'slide_type': slide_type},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=60
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP request failed: {str(e)}") from e


# 生成PPT大纲
@mcp.tool()
async def ppt_create_outline_text(
        ppt_text: str = Field(description="PPT-themeText")) -> str:
    """
    Name:
        根据用户语输入的主题文本生成大纲内容
    Description:
        根据用户输入的内容ppt_text，实时生成大纲内容，直接返回大纲文本内容。
    Args:
        ppt_text: PPT-themeText，用户输入的文本，必须参数。
    Returns:
        新的PPT-ID
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-structure'

        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                data={'text': ppt_text},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=60
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP request failed: {str(e)}") from e


# 生成PPT模板预览图
@mcp.tool()
async def ppt_create_template_cover_image(
        ppt_text: str = Field(description="PPT-themeText"),
        ppt_color: List[str] = Field(
            default_factory=lambda: ["紫色", "红色", "橙色", "黄色", "绿色", "青色", "蓝色", "粉色", "灰色"],
            description="template-color"
        ),
        ppt_style: str = Field(
            default_factory=lambda: ["科技风", "商务风", "小清新", "极简风", "中国风", "可爱卡通"],
            description="template-style"
        ),
        ppt_num: int = Field(default=4, description="Template-num")):
    """
    Name:
        根据用户语输入的主题文本生成模板，渲染封面图，并返回模板id（cover_id），用于进行替换或者生成指定模板的PPT文档。
    Description:
        根据用户输入的内容ppt_text，实时生成与渲染对应的模板，返回对应的模板ID和对应的渲染图片，支持用户指定颜色ppt_color和风格ppt_style，默认随机；也可以指定返回的数量，默认为4个。
    Args:
        ppt_text(str): PPT-themeText，用户输入主题文本，必须参数。
        ppt_color(str)：Template-color，指定生成的模板风格，可以为空，表示随机；也可以从"科技风","商务风","小清新","极简风","中国风","可爱卡通"进行执行，可选参数。
        ppt_style(str): Template-style,指定生成模板的颜色，可以为空，表示随机；也可以从"紫色","红色","橙色","黄色","绿色","青色","蓝色","粉色","灰色"进行指定，可选参数。
        ppt_num(int)：Template-num，指定生成模板数量，默认为4
    Returns:
        返回Cover-ID
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-cover'

        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                data={'title': ppt_text, "count": ppt_num, "color": ppt_color, "style": ppt_style},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=60
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP request failed: {str(e)}") from e


# 替换/更换指定模板
@mcp.tool()
async def ppt_replace_user_select_template(
        ppt_id: str = Field(description="PPT-ID"),
        cover_id: str = Field(description="cover_id")
) -> str:
    """
    Name:
        通过Cover-ID替换/更换指定模板。
    Description:
        根据工具Cover-ID和任务PPT-ID执行替换为用户指定（根据cover_id）的模板，并返回新的任务PPT-ID，可以调用query_ppt工具查询生成进度和预览URL
    Args:
        ppt_id(str): PPT-ID
        cover_id(str)：cover_id，用户指定的模板id，需要通过ppt_create_template_coverImage 进行生成。
    Returns:
        新的PPT-ID
    """

    try:
        url = API_BASE + '/mcp/ppt/ppt-create-task'

        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                data={'id': ppt_id, "cover_id": cover_id},
                headers={'Authorization': f"Bearer  {API_PPT_KEY}"},
                timeout=60
            )
            response.raise_for_status()

        if response.status_code != 200:
            raise Exception(f"API请求失败: HTTP {response.status_code}")

        return response.json()
    except httpx.HTTPError as e:
        raise Exception(f"HTTP request failed: {str(e)}") from e
