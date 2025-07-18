import logging
from typing import Optional
from mcp.client.streamable_http import streamablehttp_client
from mcp import ClientSession
import asyncio

logger = logging.getLogger(__name__)

async def process_markdown_to_doc(
    markdown_text: str,
    filename: str,
    title: str,
    author: str,
    style_file: Optional[str] = None,
    mcp_url: str = "http://wordmcp:8000/mcp"
) -> str:
    """Process Markdown text to create a Word document via MCP service.

    Args:
        markdown_text: Markdown text to process.
        filename: Path to save the Word document (e.g., 'output/report.docx').
        title: Document title for the 'Title' style.
        author: Document author for metadata.
        style_file: Optional path to JSON style file for headings and title.
        mcp_url: URL of the MCP service.

    Returns:
        str: Success message or error message.
    """
    filename = ensure_docx_extension(filename)
    logger.info(f"Processing Markdown to document: {filename}, title: {title}, style_file: {style_file}")

    async with streamablehttp_client(mcp_url) as (read_stream, write_stream, _):
        async with ClientSession(read_stream, write_stream) as session:
            # Initialize the connection
            try:
                await session.initialize()
                logger.info("Session initialized.")
            except Exception as e:
                logger.error(f"Failed to initialize session: {str(e)}")
                return f"Error: Failed to initialize MCP session: {str(e)}"

            # Create a new document
            document_input = {
                "filename": filename,
                "title": title,
                "author": author
            }
            try:
                create_result = await session.call_tool("create_document", document_input)
                logger.info(f"create_document result: {create_result}")
                print(f"document : {document_input}")
                print(f"create_result : {create_result}")
            except Exception as e:
                logger.error(f"Error in create_document: {str(e)}")
                return f"Error: Failed to create document: {str(e)}"

            # Parse Markdown and add content
            lines = markdown_text.split("\n")
            for line in lines:
                line = line.strip()
                if not line:
                    continue

                try:
                    # Handle headings (#, ##, etc.)
                    if line.startswith("#"):
                        level = len(line.split(" ")[0])  # Count # symbols for heading level
                        text = line.lstrip("# ").strip()
                        result = await session.call_tool("add_heading", {
                            "filename": filename,
                            "text": text,
                            "level": level,
                            # "style_file": style_file
                        })
                        print(f"Added heading: {text} (Level {level}) - {result}")
                        logger.info(f"Added heading: {text} (Level {level}) - {result}")

                    # Handle list items (-, *, etc.)
                    elif line.startswith("- ") or line.startswith("* "):
                        text = line[2:].strip()  # Remove list marker
                        result = await session.call_tool("add_paragraph", {
                            "filename": filename,
                            "text": f"• {text}",
                            "style": "ListBullet"
                        })
                        logger.info(f"Added list item: {text} - {result}")

                    # Handle paragraphs
                    else:
                        result = await session.call_tool("add_paragraph", {
                            "filename": filename,
                            "text": line,
                            "style": "Normal"
                        })
                        logger.info(f"Added paragraph: {line} - {result}")

                except Exception as e:
                    logger.error(f"Error adding content: {str(e)}")
                    return f"Error: Failed to add content: {str(e)}"

            return f"Document '{filename}' created successfully"

def ensure_docx_extension(filename: str) -> str:
    """Ensure the filename has a .docx extension."""
    if not filename.lower().endswith('.docx'):
        return f"{filename}.docx"
    return filename
if  __name__=="__main__":
    markdown_text = """
    # Test Heading 
    ## 测试
    Paragraph text
    
    其他内容
    测试新的问题金
    ### 其他的问题
    测AI
    #### 阿策1131
    阿发是否
    """
    result = asyncio.run(process_markdown_to_doc(markdown_text, "output/test.docx", "workflow", "Test Author"))
    print(f"Finished:{result}")
