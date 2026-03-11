import subprocess
import tempfile
import os
import asyncio
import base64
import io
import logging
import time
import json
import sys

import mcp.types as types
from mcp.server.models import InitializationOptions
from mcp.server import NotificationOptions, Server
import mcp.server.stdio
from PIL import ImageGrab

# Handle win32com import with better error handling
try:
    import win32com.client
    import pythoncom
    import win32gui
    import win32con
    WIN32_AVAILABLE = True
    print("WIN32 COM modules loaded successfully", file=sys.stderr)
except ImportError as e:
    print(f"Win32 COM not available: {e}", file=sys.stderr)
    WIN32_AVAILABLE = False
    win32com = None
    pythoncom = None
    win32gui = None
    win32con = None

try:
    from .prompt import (
        get_system_prompt,
        get_prompt_suggestions,
        get_advanced_templates,
        get_prompting_tips,
        display_help,
        format_advanced_template,
    )
except ImportError:
    from prompt import (
        get_system_prompt,
        get_prompt_suggestions,
        get_advanced_templates,
        get_prompting_tips,
        display_help,
        format_advanced_template,
    )

# Set up logging
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%dT%H:%M:%S",
)

def get_illustrator_with_retry(max_attempts=3, delay=0.3):
    if pythoncom is None or win32com is None:
        raise RuntimeError("Win32 COM not available")
    last_err: Exception = RuntimeError("Illustrator not reachable after retries")
    for attempt in range(max_attempts):
        try:
            pythoncom.CoInitialize()
            return win32com.client.GetActiveObject("Illustrator.Application")
        except Exception as e:
            last_err = e
            if attempt < max_attempts - 1:
                time.sleep(delay)
    raise last_err


server = Server("illustrator")


def _print_client_config_hint() -> None:
    """Print a ready-to-copy config snippet for MCP clients."""
    python_path = sys.executable.replace("\\", "\\\\")
    server_path = os.path.abspath(__file__).replace("\\", "\\\\")
    hint = f"""
Add this MCP config in Codex/Cursor/Claude client settings:
{{
  "mcpServers": {{
    "illustrator": {{
      "command": "{python_path}",
      "args": [
        "{server_path}"
      ]
    }}
  }}
}}
"""
    print(hint, file=sys.stderr)
    sys.stderr.flush()

@server.list_tools()
async def handle_list_tools() -> list[types.Tool]:
    logging.info("Listing available tools.")
    return [
        types.Tool(
            name="view",
            description="View a screenshot of the Adobe Illustrator window",
            inputSchema={"type": "object", "properties": {}},
        ),
        types.Tool(
            name="run",
            description="Run ExtendScript code in Illustrator",
            inputSchema={
                "type": "object",
                "properties": {
                    "code": {"type": "string", "description": "ExtendScript code to execute."}
                },
                "required": ["code"],
            },
        ),
        types.Tool(
            name="query",
            description="Query the current Illustrator document state — returns JSON with layer list, circle counts, and sample positions. Use instead of screenshots for verification.",
            inputSchema={"type": "object", "properties": {}},
        ),
        types.Tool(
            name="get_prompt_suggestions",
            description="Get categorized prompt suggestions for creating content in Illustrator",
            inputSchema={
                "type": "object",
                "properties": {
                    "category": {
                        "type": "string",
                        "description": "Optional: Filter by category (e.g., 'logos', 'illustrations', 'icons')",
                        "enum": [
                            "basic_shapes",
                            "typography",
                            "logos",
                            "illustrations", 
                            "icons",
                            "artistic",
                            "charts",
                            "print"
                        ]
                    }
                }
            },
        ),
        types.Tool(
            name="get_system_prompt",
            description="Get the system prompt template for better AI guidance when working with Illustrator",
            inputSchema={"type": "object", "properties": {}},
        ),
        types.Tool(
            name="get_prompting_tips",
            description="Get tips for creating better prompts when working with Illustrator",
            inputSchema={"type": "object", "properties": {}},
        ),
        types.Tool(
            name="get_advanced_template",
            description="Get an advanced prompt template for complex design tasks",
            inputSchema={
                "type": "object",
                "properties": {
                    "template_type": {
                        "type": "string",
                        "description": "Type of template to get",
                        "enum": ["logo_design", "illustration", "infographic", "icon_set"]
                    },
                    "parameters": {
                        "type": "object",
                        "description": "Parameters to fill in the template (varies by template type)"
                    }
                },
                "required": ["template_type"]
            },
        ),
        types.Tool(
            name="help",
            description="Display comprehensive help information for using the Illustrator MCP server",
            inputSchema={"type": "object", "properties": {}},
        ),
    ]

def _focus_illustrator():
    if win32gui is None:
        return
    results: list = []
    def cb(hwnd: int, _: object) -> None:
        if win32gui.IsWindowVisible(hwnd) and "Adobe Illustrator" in win32gui.GetWindowText(hwnd):  # type: ignore[union-attr]
            results.append(hwnd)
    win32gui.EnumWindows(cb, None)  # type: ignore[union-attr]
    if results:
        win32gui.SetForegroundWindow(results[0])  # type: ignore[union-attr]
        time.sleep(0.5)


def capture_illustrator() -> list[types.TextContent | types.ImageContent]:
    logging.info("Starting screenshot capture for Illustrator.")
    if not WIN32_AVAILABLE:
        return [types.TextContent(type="text", text="Win32 COM not available. Please install pywin32 and restart the server.")]

    try:
        _focus_illustrator()
        screenshot = ImageGrab.grab()
        buffer = io.BytesIO()
        screenshot.save(buffer, format="JPEG", quality=50, optimize=True)
        screenshot_data = base64.b64encode(buffer.getvalue()).decode("utf-8")
        logging.info("Screenshot captured successfully.")
        return [types.ImageContent(type="image", mimeType="image/jpeg", data=screenshot_data)]
    except Exception as e:
        logging.error(f"Failed to capture screenshot: {str(e)}")
        return [types.TextContent(type="text", text=f"Failed to capture screenshot: {str(e)}")]

def run_illustrator_script(code: str) -> list[types.TextContent]:
    logging.info("Running ExtendScript code in Illustrator using COM.")
    if not WIN32_AVAILABLE:
        return [types.TextContent(type="text", text="Win32 COM not available. Please install pywin32 and restart the server.")]

    try:
        with tempfile.NamedTemporaryFile(suffix=".jsx", delete=False) as jsx_file:
            jsx_file.write(code.encode("utf-8"))
            jsx_file_path = jsx_file.name
        logging.debug(f"ExtendScript saved to: {jsx_file_path}")
        illustrator = get_illustrator_with_retry()
        result = illustrator.DoJavaScriptFile(jsx_file_path)
        logging.info("ExtendScript executed successfully.")
        os.unlink(jsx_file_path)
        logging.debug("Temporary ExtendScript file removed.")
        result_text = str(result) if result is not None else "Script executed successfully (no return value)"
        return [types.TextContent(type="text", text=result_text)]
    except Exception as e:
        logging.error(f"Failed to execute script: {str(e)}")
        return [types.TextContent(type="text", text=f"Failed to execute script: {str(e)}")]

def query_illustrator_state() -> list[types.TextContent]:
    logging.info("Querying Illustrator document state.")
    if not WIN32_AVAILABLE:
        return [types.TextContent(type="text", text="Win32 COM not available.")]

    jsx = r"""
(function() {
    var doc;
    try { doc = app.activeDocument; } catch(e) { return '{"error":"No document open"}'; }

    var ab = doc.artboards[0].artboardRect;
    var artboard = [ab[0], ab[1], ab[2], ab[3]];

    var layers = [];
    for (var i = 0; i < doc.layers.length; i++) {
        var l = doc.layers[i];
        layers.push({ name: l.name, visible: l.visible, locked: l.locked });
    }

    // Find TPZs layer using only pathItems (never pageItems)
    var tpzLayer = null;
    for (var li = 0; li < doc.layers.length; li++) {
        if (doc.layers[li].name === 'TPZs') { tpzLayer = doc.layers[li]; break; }
    }

    var tpzCircles = [];
    var tpzCount = 0;
    var trunkCount = 0;

    if (tpzLayer) {
        // Count direct path items that look like circles (TPZ layer)
        var paths = tpzLayer.pathItems;
        for (var pi = 0; pi < paths.length; pi++) {
            var item = paths[pi];
            // Only consider direct children of TPZs layer
            try { if (item.parent.name !== 'TPZs') continue; } catch(e) { continue; }
            var gb = item.geometricBounds;
            var w = Math.abs(gb[2] - gb[0]);
            var h = Math.abs(gb[1] - gb[3]);
            if (w <= 0 || h <= 0) continue;
            var isCircle = Math.abs(w - h) / Math.max(w, h) < 0.15 && item.pathPoints.length >= 4 && item.pathPoints.length <= 5;
            if (!isCircle) continue;
            tpzCount++;
            if (tpzCircles.length < 5) {
                var cx = (gb[0] + gb[2]) / 2;
                var cy = (gb[1] + gb[3]) / 2;
                var colorName = "unknown";
                try {
                    var sc = item.strokeColor;
                    if (sc.typename === 'CMYKColor') {
                        colorName = sc.cyan >= 60 ? "green" : "orange";
                    }
                } catch(e) {}
                tpzCircles.push({ cx: Math.round(cx), cy: Math.round(cy), diameter: Math.round(w), color: colorName, dashed: item.strokeDashes && item.strokeDashes.length > 0 });
            }
        }

        // Count trunk circles in Trunks sublayer
        for (var sli = 0; sli < tpzLayer.layers.length; sli++) {
            var sub = tpzLayer.layers[sli];
            if (sub.name === 'Trunks') {
                var tpaths = sub.pathItems;
                for (var tpi = 0; tpi < tpaths.length; tpi++) {
                    var titem = tpaths[tpi];
                    var tgb = titem.geometricBounds;
                    var tw = Math.abs(tgb[2] - tgb[0]);
                    var th = Math.abs(tgb[1] - tgb[3]);
                    if (tw <= 0 || th <= 0) continue;
                    if (Math.abs(tw - th) / Math.max(tw, th) < 0.15) trunkCount++;
                }
                break;
            }
        }
    }

    // Manual JSON serialization — ExtendScript has no native JSON object
    var layersJson = '[';
    for (var lj = 0; lj < layers.length; lj++) {
        if (lj > 0) layersJson += ',';
        var ll = layers[lj];
        layersJson += '{"name":"' + ll.name.replace(/\\/g,'\\\\').replace(/"/g,'\\"') + '"' +
            ',"visible":' + (ll.visible ? 'true' : 'false') +
            ',"locked":'  + (ll.locked  ? 'true' : 'false') + '}';
    }
    layersJson += ']';

    var circlesJson = '[';
    for (var cj = 0; cj < tpzCircles.length; cj++) {
        if (cj > 0) circlesJson += ',';
        var cc = tpzCircles[cj];
        circlesJson += '{"cx":' + cc.cx + ',"cy":' + cc.cy + ',"diameter":' + cc.diameter +
            ',"color":"' + cc.color + '","dashed":' + (cc.dashed ? 'true' : 'false') + '}';
    }
    circlesJson += ']';

    return '{"doc":"' + doc.name.replace(/\\/g,'\\\\').replace(/"/g,'\\"') + '"' +
        ',"artboard":[' + artboard[0] + ',' + artboard[1] + ',' + artboard[2] + ',' + artboard[3] + ']' +
        ',"layers":' + layersJson +
        ',"tpz_circles":' + circlesJson +
        ',"tpz_count":' + tpzCount +
        ',"trunk_count":' + trunkCount +
        ',"error":null}';
})();
"""
    try:
        illustrator = get_illustrator_with_retry()
        with tempfile.NamedTemporaryFile(suffix=".jsx", delete=False) as jsx_file:
            jsx_file.write(jsx.encode("utf-8"))
            jsx_file_path = jsx_file.name
        result = illustrator.DoJavaScriptFile(jsx_file_path)
        os.unlink(jsx_file_path)
        result_text = str(result) if result is not None else '{"error": "No result returned"}'
        return [types.TextContent(type="text", text=result_text)]
    except Exception as e:
        logging.error(f"Failed to query state: {str(e)}")
        return [types.TextContent(type="text", text=f'{{"error": "{str(e)}"}}')]


@server.call_tool()
async def handle_call_tool(name: str, arguments: dict | None):
    logging.info(f"Received tool call: {name} with arguments: {arguments}")
    
    if name == "view":
        return capture_illustrator()
    
    elif name == "run":
        if not arguments or "code" not in arguments:
            logging.warning("No code provided for run tool.")
            return [types.TextContent(type="text", text="No code provided")]
        return run_illustrator_script(arguments["code"])
    
    elif name == "query":
        return query_illustrator_state()

    elif name == "get_prompt_suggestions":
        try:
            suggestions = get_prompt_suggestions()
            category = arguments.get("category") if arguments else None
            
            if category:
                # Filter by category
                category_map = {
                    "basic_shapes": "🎨 Basic Shapes & Geometry",
                    "typography": "📝 Typography & Text", 
                    "logos": "🏢 Logos & Branding",
                    "illustrations": "🌆 Illustrations & Scenes",
                    "icons": "🎭 Icons & UI Elements",
                    "artistic": "🎨 Artistic & Creative",
                    "charts": "📊 Charts & Infographics",
                    "print": "🏷️ Print & Layout"
                }
                
                full_category = category_map.get(category)
                if full_category and full_category in suggestions:
                    filtered_suggestions = {full_category: suggestions[full_category]}
                    result_text = f"**{full_category}**\n\n"
                    for suggestion in suggestions[full_category]:
                        result_text += f"• {suggestion}\n"
                else:
                    result_text = f"Category '{category}' not found. Available categories: {list(category_map.keys())}"
            else:
                # Return all suggestions
                result_text = "# 🎨 Illustrator Prompt Suggestions\n\n"
                for category, prompts in suggestions.items():
                    result_text += f"## {category}\n\n"
                    for prompt in prompts:
                        result_text += f"• {prompt}\n"
                    result_text += "\n"
            
            return [types.TextContent(type="text", text=result_text)]
        except Exception as e:
            logging.error(f"Error getting prompt suggestions: {str(e)}")
            return [types.TextContent(type="text", text=f"Error: {str(e)}")]
    
    elif name == "get_system_prompt":
        try:
            system_prompt = get_system_prompt()
            return [types.TextContent(type="text", text=system_prompt)]
        except Exception as e:
            logging.error(f"Error getting system prompt: {str(e)}")
            return [types.TextContent(type="text", text=f"Error: {str(e)}")]
    
    elif name == "get_prompting_tips":
        try:
            tips = get_prompting_tips()
            result_text = "# 💡 Prompting Tips for Adobe Illustrator\n\n"
            for tip in tips:
                result_text += f"{tip}\n"
            return [types.TextContent(type="text", text=result_text)]
        except Exception as e:
            logging.error(f"Error getting prompting tips: {str(e)}")
            return [types.TextContent(type="text", text=f"Error: {str(e)}")]
    
    elif name == "get_advanced_template":
        try:
            template_type = arguments.get("template_type") if arguments else None
            parameters = arguments.get("parameters", {}) if arguments else {}
            
            if not template_type:
                return [types.TextContent(type="text", text="Template type is required")]
            
            templates = get_advanced_templates()
            if template_type in templates:
                if parameters:
                    # Try to format with parameters
                    try:
                        formatted_template = format_advanced_template(template_type, **parameters)
                        return [types.TextContent(type="text", text=formatted_template)]
                    except KeyError as e:
                        # Missing parameters, return template with placeholders
                        template = templates[template_type]
                        result_text = f"**{template_type.replace('_', ' ').title()} Template:**\n\n{template}\n\n"
                        result_text += f"**Missing parameter:** {str(e)}\n"
                        result_text += "Please provide the required parameters to fill in the template."
                        return [types.TextContent(type="text", text=result_text)]
                else:
                    # Return template with placeholders
                    template = templates[template_type]
                    result_text = f"**{template_type.replace('_', ' ').title()} Template:**\n\n{template}"
                    return [types.TextContent(type="text", text=result_text)]
            else:
                available_templates = list(templates.keys())
                return [types.TextContent(type="text", text=f"Template '{template_type}' not found. Available templates: {available_templates}")]
        except Exception as e:
            logging.error(f"Error getting advanced template: {str(e)}")
            return [types.TextContent(type="text", text=f"Error: {str(e)}")]
    
    elif name == "help":
        try:
            help_text = display_help()
            return [types.TextContent(type="text", text=help_text)]
        except Exception as e:
            logging.error(f"Error displaying help: {str(e)}")
            return [types.TextContent(type="text", text=f"Error: {str(e)}")]
    
    else:
        error_msg = f"Unknown tool: {name}"
        logging.error(error_msg)
        raise ValueError(error_msg)

async def main():
    try:
        print("Initializing MCP server for Illustrator...", file=sys.stderr)
        sys.stderr.flush()
        logging.info("Initializing MCP server for Illustrator.")
        
        async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
            print("Server streams established, starting server...", file=sys.stderr)
            sys.stderr.flush()
            _print_client_config_hint()
            
            await server.run(
                read_stream,
                write_stream,
                InitializationOptions(
                    server_name="illustrator",
                    server_version="0.1.0",
                    capabilities=server.get_capabilities(
                        notification_options=NotificationOptions(),
                        experimental_capabilities={},
                    ),
                ),
            )
            print("Server finished running", file=sys.stderr)
            sys.stderr.flush()
    except Exception as e:
        print(f"Error in main: {e}", file=sys.stderr)
        sys.stderr.flush()
        import traceback
        traceback.print_exc(file=sys.stderr)
        sys.stderr.flush()
        raise

if __name__ == "__main__":
    try:
        print("Starting the main event loop...", file=sys.stderr)
        logging.info("Starting the main event loop.")
        asyncio.run(main())
    except Exception as e:
        print(f"Error starting server: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)
        sys.exit(1)
