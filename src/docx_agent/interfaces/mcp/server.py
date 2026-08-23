"""
Standard Model Context Protocol (MCP) Stdio Server for docx-agent platform.
Exposes high-level, format-preserving DOCX tools to AI coding agents.
"""

import sys
import json
import traceback
from typing import Dict, Any, List, Optional
from docx_agent.agent import DocumentAgent
from docx_agent.core.exceptions import DocxAgentError
from docx_agent.utils.logging import logger

# ----------------------------------------------------------------------
# MCP TOOL DEFINITIONS AND SCHEMAS
# ----------------------------------------------------------------------

MCP_TOOLS = [
    {
        "name": "docx_inspect",
        "description": "Inspects structural summary, total elements, page geometries, margins, and outline of a .docx file.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file": {"type": "string", "description": "Absolute or relative path to the .docx file"}
            },
            "required": ["file"],
        },
    },
    {
        "name": "docx_read",
        "description": "Reads document paragraphs with stable element IDs, styles, and optional run-level formatting metadata.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file": {"type": "string", "description": "Path to .docx file"},
                "start": {"type": "integer", "default": 0, "description": "Start paragraph index"},
                "end": {"type": "integer", "description": "End paragraph index (exclusive)"},
                "include_runs": {"type": "boolean", "default": False, "description": "Include run styling details"},
                "show_tables": {"type": "boolean", "default": False, "description": "Include table content"},
            },
            "required": ["file"],
        },
    },
    {
        "name": "docx_outline",
        "description": "Extracts document heading outline hierarchy (Title, Heading 1, Heading 2, Heading 3...).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file": {"type": "string", "description": "Path to .docx file"}
            },
            "required": ["file"],
        },
    },
    {
        "name": "docx_find",
        "description": "Finds paragraphs matching text substring, regular expression, or style name.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file": {"type": "string", "description": "Path to .docx file"},
                "text": {"type": "string", "description": "Substring to search for"},
                "regex": {"type": "string", "description": "Regular expression pattern"},
                "style": {"type": "string", "description": "Style name filter"},
            },
            "required": ["file"],
        },
    },
    {
        "name": "docx_capabilities",
        "description": "Inspects and reports supported features, warnings, or unsupported embedded objects in a document.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file": {"type": "string", "description": "Path to .docx file"}
            },
            "required": ["file"],
        },
    },
    {
        "name": "docx_replace",
        "description": "Replaces text while strictly preserving surrounding run formatting (bold, italic, colors, fonts, hyperlinks).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file": {"type": "string", "description": "Path to .docx file"},
                "target": {"type": "string", "description": "Target string to replace"},
                "replacement": {"type": "string", "description": "Replacement string"},
                "scope": {"type": "string", "description": "Optional element ID (e.g. 'p_0012') or selector"},
                "output": {"type": "string", "description": "Output path (defaults to in-place)"},
            },
            "required": ["file", "target", "replacement"],
        },
    },
    {
        "name": "docx_insert",
        "description": "Inserts a new paragraph before or after a target element.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file": {"type": "string", "description": "Path to .docx file"},
                "text": {"type": "string", "description": "Paragraph text to insert"},
                "target": {"type": "string", "description": "Target element ID (e.g. 'p_0012')"},
                "position": {"type": "string", "enum": ["before", "after"], "default": "after"},
                "style": {"type": "string", "description": "Style name"},
                "output": {"type": "string", "description": "Output path"},
            },
            "required": ["file", "text", "target"],
        },
    },
    {
        "name": "docx_format_text",
        "description": "Applies font family, font size, bold, italic, color, and character formatting to target paragraph(s).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file": {"type": "string", "description": "Path to .docx file"},
                "target": {"type": "string", "description": "Element ID or selector"},
                "font_name": {"type": "string", "description": "Font name (e.g. 'Times New Roman')"},
                "font_size_pt": {"type": "number", "description": "Font size in points"},
                "bold": {"type": "boolean"},
                "italic": {"type": "boolean"},
                "underline": {"type": "boolean"},
                "color_rgb": {"type": "string", "description": "Hex color string (e.g. '003366')"},
                "output": {"type": "string", "description": "Output path"},
            },
            "required": ["file", "target"],
        },
    },
    {
        "name": "docx_format_paragraph",
        "description": "Applies paragraph alignment, line spacing, margins, and space before/after.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file": {"type": "string", "description": "Path to .docx file"},
                "target": {"type": "string", "description": "Element ID or selector"},
                "alignment": {"type": "string", "enum": ["left", "center", "right", "justify"]},
                "line_spacing": {"type": "number", "description": "Line spacing multiplier (e.g. 1.5)"},
                "space_before_pt": {"type": "number", "description": "Space before in points"},
                "space_after_pt": {"type": "number", "description": "Space after in points"},
                "first_line_indent_cm": {"type": "number", "description": "First-line indent in cm"},
                "output": {"type": "string", "description": "Output path"},
            },
            "required": ["file", "target"],
        },
    },
    {
        "name": "docx_preset",
        "description": "Applies complete document formatting preset ('academic-vn', 'ieee', 'apa', 'technical-report').",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file": {"type": "string", "description": "Path to .docx file"},
                "name": {"type": "string", "default": "academic-vn", "description": "Preset name or JSON path"},
                "output": {"type": "string", "description": "Output path"},
            },
            "required": ["file"],
        },
    },
    {
        "name": "docx_table",
        "description": "Creates a structured table with repeating headers, custom borders, and cell formatting.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file": {"type": "string", "description": "Path to .docx file"},
                "rows": {"type": "integer", "description": "Number of rows"},
                "cols": {"type": "integer", "description": "Number of columns"},
                "data": {"type": "array", "items": {"type": "array", "items": {"type": "string"}}},
                "style": {"type": "string", "default": "Table Grid"},
                "output": {"type": "string", "description": "Output path"},
            },
            "required": ["file", "rows", "cols"],
        },
    },
    {
        "name": "docx_image",
        "description": "Inserts an image into the document with scaling, alignment, and optional caption.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file": {"type": "string", "description": "Path to .docx file"},
                "image_path": {"type": "string", "description": "Path to image file (PNG, JPG)"},
                "target": {"type": "string", "description": "Target paragraph ID"},
                "caption": {"type": "string", "description": "Caption text"},
                "width_cm": {"type": "number", "description": "Width in centimeters"},
                "output": {"type": "string", "description": "Output path"},
            },
            "required": ["file", "image_path"],
        },
    },
    {
        "name": "docx_verify",
        "description": "Independently validates file integrity and verifies typography against expected criteria.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file": {"type": "string", "description": "Path to .docx file"},
                "expected_font": {"type": "string", "description": "Expected body font name"},
                "expected_size": {"type": "number", "description": "Expected body size in pt"},
                "expected_line_spacing": {"type": "number", "description": "Expected line spacing multiplier"},
                "expected_alignment": {"type": "string", "description": "Expected alignment"},
            },
            "required": ["file"],
        },
    },
    {
        "name": "docx_diff",
        "description": "Generates semantic structural and textual diff between two DOCX revisions.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "before": {"type": "string", "description": "Path to before .docx file"},
                "after": {"type": "string", "description": "Path to after .docx file"},
            },
            "required": ["before", "after"],
        },
    },
    {
        "name": "docx_apply_plan",
        "description": "Applies a multi-step batch operation plan atomically with pre-validation and rollback.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "plan": {"type": "object", "description": "Batch plan object or dict"},
                "output": {"type": "string", "description": "Output path"},
            },
            "required": ["plan"],
        },
    },
]


def handle_tool_call(name: str, args: Dict[str, Any]) -> Dict[str, Any]:
    file_p = args.get("file")
    out_p = args.get("output")

    if name == "docx_inspect":
        agent = DocumentAgent(file_p)
        return agent.inspect().model_dump()

    elif name == "docx_read":
        agent = DocumentAgent(file_p)
        paras = agent.read(
            start=args.get("start", 0),
            end=args.get("end"),
            include_runs=args.get("include_runs", False),
        )
        res: Dict[str, Any] = {"paragraphs": [p.model_dump(exclude_none=True) for p in paras]}
        if args.get("show_tables"):
            res["tables"] = [t.model_dump(exclude_none=True) for t in agent.model.get_tables()]
        return res

    elif name == "docx_outline":
        agent = DocumentAgent(file_p)
        return {"outline": agent.outline()}

    elif name == "docx_find":
        agent = DocumentAgent(file_p)
        return {"matches": agent.find(text=args.get("text"), regex=args.get("regex"), style=args.get("style"))}

    elif name == "docx_capabilities":
        agent = DocumentAgent(file_p)
        return agent.capabilities()

    elif name == "docx_replace":
        agent = DocumentAgent(file_p)
        n = agent.replace(
            target=args["target"],
            replacement=args["replacement"],
            scope=args.get("scope"),
        )
        saved = agent.save(output_path=out_p)
        return {"success": True, "replaced_count": n, "document": saved}

    elif name == "docx_insert":
        agent = DocumentAgent(file_p)
        new_id = agent.insert(
            text=args["text"],
            target=args["target"],
            position=args.get("position", "after"),
            style=args.get("style"),
        )
        saved = agent.save(output_path=out_p)
        return {"success": True, "element_id": new_id, "document": saved}

    elif name == "docx_format_text":
        agent = DocumentAgent(file_p)
        n = agent.format_text(
            target=args["target"],
            font_name=args.get("font_name"),
            font_size_pt=args.get("font_size_pt"),
            bold=args.get("bold"),
            italic=args.get("italic"),
            underline=args.get("underline"),
            color_rgb=args.get("color_rgb"),
        )
        saved = agent.save(output_path=out_p)
        return {"success": True, "paragraphs_formatted": n, "document": saved}

    elif name == "docx_format_paragraph":
        agent = DocumentAgent(file_p)
        n = agent.format_paragraph(
            target=args["target"],
            alignment=args.get("alignment"),
            line_spacing=args.get("line_spacing"),
            space_before_pt=args.get("space_before_pt"),
            space_after_pt=args.get("space_after_pt"),
            first_line_indent_cm=args.get("first_line_indent_cm"),
        )
        saved = agent.save(output_path=out_p)
        return {"success": True, "paragraphs_formatted": n, "document": saved}

    elif name == "docx_preset":
        agent = DocumentAgent(file_p)
        details = agent.apply_preset(args.get("name", "academic-vn"))
        saved = agent.save(output_path=out_p)
        return {"success": True, "preset_details": details, "document": saved}

    elif name == "docx_table":
        agent = DocumentAgent(file_p)
        tid = agent.create_table(
            rows=args["rows"],
            cols=args["cols"],
            data=args.get("data"),
            style=args.get("style", "Table Grid"),
        )
        saved = agent.save(output_path=out_p)
        return {"success": True, "table_id": tid, "document": saved}

    elif name == "docx_image":
        agent = DocumentAgent(file_p)
        pid = agent.insert_image(
            image_path=args["image_path"],
            target=args.get("target"),
            caption=args.get("caption"),
            width_cm=args.get("width_cm"),
        )
        saved = agent.save(output_path=out_p)
        return {"success": True, "paragraph_id": pid, "document": saved}

    elif name == "docx_verify":
        agent = DocumentAgent(file_p)
        return agent.verify(
            expected_font=args.get("expected_font"),
            expected_size_pt=args.get("expected_size"),
            expected_line_spacing=args.get("expected_line_spacing"),
            expected_alignment=args.get("expected_alignment"),
        )

    elif name == "docx_diff":
        agent = DocumentAgent(args["before"])
        diff_rep = agent.diff(args["after"])
        return diff_rep.model_dump()

    elif name == "docx_apply_plan":
        plan_data = args["plan"]
        doc_path = plan_data.get("document")
        agent = DocumentAgent(doc_path)
        plan_res = agent.apply_plan(plan_data)
        saved = agent.save(output_path=out_p)
        return {"success": True, "plan_result": plan_res, "document": saved}

    else:
        raise ValueError(f"Unknown MCP tool: {name}")


def main():
    """Main JSON-RPC stdio loop for MCP clients."""
    if sys.stdout.encoding != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")
    if sys.stdin.encoding != "utf-8":
        sys.stdin.reconfigure(encoding="utf-8")

    logger.info("docx-agent MCP Server started on stdio.")

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue

        try:
            request = json.loads(line)
            req_id = request.get("id")
            method = request.get("method")
            params = request.get("params", {})

            if method == "initialize":
                response = {
                    "jsonrpc": "2.0",
                    "id": req_id,
                    "result": {
                        "protocolVersion": "2024-11-05",
                        "capabilities": {"tools": {}},
                        "serverInfo": {"name": "docx-agent-mcp", "version": "1.0.0"},
                    },
                }
            elif method == "tools/list":
                response = {
                    "jsonrpc": "2.0",
                    "id": req_id,
                    "result": {"tools": MCP_TOOLS},
                }
            elif method == "tools/call":
                tool_name = params.get("name")
                tool_args = params.get("arguments", {})
                result_payload = handle_tool_call(tool_name, tool_args)
                response = {
                    "jsonrpc": "2.0",
                    "id": req_id,
                    "result": {
                        "content": [
                            {"type": "text", "text": json.dumps(result_payload, indent=2, ensure_ascii=False)}
                        ]
                    },
                }
            else:
                response = {
                    "jsonrpc": "2.0",
                    "id": req_id,
                    "error": {"code": -32601, "message": f"Method not found: {method}"},
                }

        except Exception as e:
            req_id = request.get("id") if "request" in locals() and isinstance(request, dict) else None
            response = {
                "jsonrpc": "2.0",
                "id": req_id,
                "error": {
                    "code": -32000,
                    "message": str(e),
                    "data": {"traceback": traceback.format_exc()},
                },
            }

        sys.stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
        sys.stdout.flush()


if __name__ == "__main__":
    main()
