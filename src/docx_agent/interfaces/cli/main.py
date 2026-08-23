"""
Main CLI Interface for docx-agent platform.
Supports rich human output and machine-readable JSON for AI coding agents.
"""

import sys
import json
from pathlib import Path
from typing import Optional, List, Any
import typer
from rich.console import Console
from rich.table import Table as RichTable
from rich.panel import Panel

from docx_agent.agent import DocumentAgent
from docx_agent.core.exceptions import DocxAgentError
from docx_agent.operations.markdown import MarkdownToDocxConverter
from docx_agent.utils.logging import logger

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
if hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

app = typer.Typer(
    name="docx-agent",
    help="Universal, Agent-Native Microsoft Word (.docx) Engine",
    add_completion=False,
)
err_console = Console(stderr=True)


def output_result(data: Any, as_json: bool) -> None:
    if as_json:
        if hasattr(data, "model_dump"):
            data_dict = data.model_dump(exclude_none=True)
        elif hasattr(data, "to_dict"):
            data_dict = data.to_dict()
        else:
            data_dict = data
        sys.stdout.write(json.dumps(data_dict, indent=2, ensure_ascii=False) + "\n")
        sys.stdout.flush()
    else:
        if isinstance(data, dict) or isinstance(data, list):
            sys.stdout.write(json.dumps(data, indent=2, ensure_ascii=False) + "\n")
        else:
            print(data)


# ----------------------------------------------------------------------
# INSPECT & DISCOVERY COMMANDS
# ----------------------------------------------------------------------

@app.command("inspect")
def cmd_inspect(
    file: str = typer.Argument(..., help="Path to .docx file"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Inspects document metadata, sections, and heading outline."""
    try:
        agent = DocumentAgent(file)
        summary = agent.inspect()
        if as_json:
            output_result(summary, as_json=True)
        else:
            err_console.print(f"[bold cyan]=== Document Summary: {Path(file).name} ===[/bold cyan]")
            err_console.print(f"Total Paragraphs: {summary.paragraphs_count}")
            err_console.print(f"Total Tables: {summary.tables_count}")
            err_console.print(f"Total Sections: {summary.sections_count}")
            if summary.sections:
                s0 = summary.sections[0]
                err_console.print(f"Page Size: {s0.page_width_cm}cm x {s0.page_height_cm}cm ({s0.orientation})")
                err_console.print(f"Margins (T, B, L, R): {s0.margin_top_cm}cm, {s0.margin_bottom_cm}cm, {s0.margin_left_cm}cm, {s0.margin_right_cm}cm")

            err_console.print("\n[bold]--- Headings Outline ---[/bold]")
            for h in summary.headings_outline:
                indent = "  " * (h["level"] if h["level"] > 0 else 0)
                err_console.print(f"[{h['id']}] {indent}{h['style']}: {h['text']}")
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("read")
def cmd_read(
    file: str = typer.Argument(..., help="Path to .docx file"),
    start: int = typer.Option(0, "--start", help="Start paragraph index"),
    end: Optional[int] = typer.Option(None, "--end", help="End paragraph index"),
    runs: bool = typer.Option(False, "--runs", help="Include run-level styling metadata"),
    show_tables: bool = typer.Option(False, "--show-tables", help="Include table text"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Reads document content with element IDs and paragraph indices."""
    try:
        agent = DocumentAgent(file)
        paras = agent.read(start=start, end=end, include_runs=runs)
        if as_json:
            res = {"paragraphs": [p.model_dump(exclude_none=True) for p in paras]}
            if show_tables:
                res["tables"] = [t.model_dump(exclude_none=True) for t in agent.model.get_tables()]
            output_result(res, as_json=True)
        else:
            for p in paras:
                st = f" ({p.style})" if p.style != "Normal" else ""
                print(f"[{p.id}]{st} {p.text}")
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("outline")
def cmd_outline(
    file: str = typer.Argument(..., help="Path to .docx file"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Extracts document heading outline hierarchy."""
    try:
        agent = DocumentAgent(file)
        out = agent.outline()
        output_result(out, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("capabilities")
def cmd_capabilities(
    file: str = typer.Argument(..., help="Path to .docx file"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Reports supported features and warnings for a document."""
    try:
        agent = DocumentAgent(file)
        cap = agent.capabilities()
        output_result(cap, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("find")
def cmd_find(
    file: str = typer.Argument(..., help="Path to .docx file"),
    text: Optional[str] = typer.Option(None, "--text", help="Substring text query"),
    regex: Optional[str] = typer.Option(None, "--regex", help="Regular expression query"),
    style: Optional[str] = typer.Option(None, "--style", help="Style name query"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Finds paragraphs matching text, regex, or style selector."""
    try:
        agent = DocumentAgent(file)
        results = agent.find(text=text, regex=regex, style=style)
        output_result(results, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


# ----------------------------------------------------------------------
# CONTENT MUTATION COMMANDS
# ----------------------------------------------------------------------

@app.command("replace")
def cmd_replace(
    file: str = typer.Argument(..., help="Path to .docx file"),
    target: str = typer.Option(..., "--target", help="Target string to find"),
    replacement: str = typer.Option(..., "--replace", "--replacement", help="Replacement string"),
    scope: Optional[str] = typer.Option(None, "--scope", help="Element ID or selector scope"),
    count: Optional[int] = typer.Option(None, "--count", help="Maximum occurrences to replace"),
    output: Optional[str] = typer.Option(None, "--output", help="Output file path (default: in-place)"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Replaces text while strictly preserving run formatting."""
    try:
        agent = DocumentAgent(file)
        n = agent.replace(target=target, replacement=replacement, scope=scope, count=count)
        out_p = agent.save(output_path=output)
        res = {"success": True, "operation": "replace", "document": out_p, "replaced_count": n}
        output_result(res, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("insert")
def cmd_insert(
    file: str = typer.Argument(..., help="Path to .docx file"),
    text: str = typer.Option(..., "--text", help="Paragraph text to insert"),
    target: str = typer.Option(..., "--target", help="Reference element ID or selector"),
    position: str = typer.Option("after", "--position", help="Insert 'before' or 'after' target"),
    style: Optional[str] = typer.Option(None, "--style", help="Style name"),
    output: Optional[str] = typer.Option(None, "--output", help="Output file path"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Inserts a paragraph before or after a target element."""
    try:
        agent = DocumentAgent(file)
        new_id = agent.insert(text=text, target=target, position=position, style=style)
        out_p = agent.save(output_path=output)
        res = {"success": True, "operation": "insert", "document": out_p, "element_id": new_id}
        output_result(res, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("append-para")
def cmd_append_para(
    file: str = typer.Argument(..., help="Path to .docx file"),
    text: str = typer.Option(..., "--text", help="Paragraph text"),
    style: Optional[str] = typer.Option("Normal", "--style", help="Style name"),
    level: Optional[int] = typer.Option(None, "--level", help="Heading level (1-9)"),
    output: Optional[str] = typer.Option(None, "--output", help="Output file path"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Appends a new paragraph or heading to the document."""
    try:
        agent = DocumentAgent(file if Path(file).exists() else None)
        new_id = agent.append(text=text, style=style, heading_level=level)
        out_p = agent.save(output_path=output or file)
        res = {"success": True, "operation": "append_para", "document": out_p, "element_id": new_id}
        output_result(res, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("edit-para")
def cmd_edit_para(
    file: str = typer.Argument(..., help="Path to .docx file"),
    index: int = typer.Option(..., "--index", help="0-based paragraph index"),
    text: str = typer.Option(..., "--text", help="New paragraph text"),
    output: Optional[str] = typer.Option(None, "--output", help="Output file path"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Edits content of a specific paragraph by index (backward-compatible)."""
    try:
        agent = DocumentAgent(file)
        p_target = f"idx:{index}"
        paras = agent.resolver.resolve_paragraphs(p_target, single=True)
        old_text = paras[0].text
        agent.replace(target=old_text, replacement=text, scope=p_target)
        out_p = agent.save(output_path=output)
        res = {"success": True, "operation": "edit_para", "document": out_p, "index": index}
        output_result(res, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("delete")
def cmd_delete(
    file: str = typer.Argument(..., help="Path to .docx file"),
    target: str = typer.Option(..., "--target", help="Element ID (e.g. p_0012, tbl_0001)"),
    output: Optional[str] = typer.Option(None, "--output", help="Output file path"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Deletes an element from the document."""
    try:
        agent = DocumentAgent(file)
        agent.delete(target)
        out_p = agent.save(output_path=output)
        res = {"success": True, "operation": "delete", "document": out_p, "target": target}
        output_result(res, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


# ----------------------------------------------------------------------
# FORMATTING & STYLES COMMANDS
# ----------------------------------------------------------------------

@app.command("format-text")
def cmd_format_text(
    file: str = typer.Argument(..., help="Path to .docx file"),
    target: str = typer.Option(..., "--target", help="Element ID or selector"),
    font_name: Optional[str] = typer.Option(None, "--font-name", help="Font family (e.g. 'Times New Roman')"),
    font_size_pt: Optional[float] = typer.Option(None, "--font-size-pt", help="Font size in points"),
    bold: Optional[bool] = typer.Option(None, "--bold/--no-bold"),
    italic: Optional[bool] = typer.Option(None, "--italic/--no-italic"),
    color_rgb: Optional[str] = typer.Option(None, "--color-rgb", help="Hex color string (e.g. '003366')"),
    output: Optional[str] = typer.Option(None, "--output", help="Output file path"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Formats font and character properties on target paragraph(s)."""
    try:
        agent = DocumentAgent(file)
        n = agent.format_text(
            target=target,
            font_name=font_name,
            font_size_pt=font_size_pt,
            bold=bold,
            italic=italic,
            color_rgb=color_rgb,
        )
        out_p = agent.save(output_path=output)
        res = {"success": True, "operation": "format_text", "document": out_p, "paragraphs_formatted": n}
        output_result(res, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("format-paragraph")
def cmd_format_paragraph(
    file: str = typer.Argument(..., help="Path to .docx file"),
    target: str = typer.Option(..., "--target", help="Element ID or selector"),
    alignment: Optional[str] = typer.Option(None, "--alignment", help="left, center, right, justify"),
    line_spacing: Optional[float] = typer.Option(None, "--line-spacing", help="1.0, 1.15, 1.5, 2.0"),
    space_before_pt: Optional[float] = typer.Option(None, "--space-before-pt", help="Points before"),
    space_after_pt: Optional[float] = typer.Option(None, "--space-after-pt", help="Points after"),
    first_line_indent_cm: Optional[float] = typer.Option(None, "--first-line-indent-cm", help="Cm indent"),
    output: Optional[str] = typer.Option(None, "--output", help="Output file path"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Formats paragraph layout, spacing, indentation, and alignment."""
    try:
        agent = DocumentAgent(file)
        n = agent.format_paragraph(
            target=target,
            alignment=alignment,
            line_spacing=line_spacing,
            space_before_pt=space_before_pt,
            space_after_pt=space_after_pt,
            first_line_indent_cm=first_line_indent_cm,
        )
        out_p = agent.save(output_path=output)
        res = {"success": True, "operation": "format_paragraph", "document": out_p, "paragraphs_formatted": n}
        output_result(res, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("preset")
def cmd_preset(
    file: str = typer.Argument(..., help="Path to .docx file"),
    name: str = typer.Option("academic-vn", "--name", help="Preset name ('academic-vn', 'ieee', 'apa', etc.)"),
    output: Optional[str] = typer.Option(None, "--output", help="Output file path"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Applies institutional formatting preset to document."""
    try:
        agent = DocumentAgent(file)
        rep = agent.apply_preset(name)
        out_p = agent.save(output_path=output)
        res = {"success": True, "operation": "apply_preset", "document": out_p, "details": rep}
        output_result(res, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


# ----------------------------------------------------------------------
# VERIFICATION, DIFF & PLAN COMMANDS
# ----------------------------------------------------------------------

@app.command("verify")
def cmd_verify(
    file: str = typer.Argument(..., help="Path to .docx file"),
    expected_font: Optional[str] = typer.Option(None, "--expected-font", help="Expected body font"),
    expected_size: Optional[float] = typer.Option(None, "--expected-size", help="Expected body size (pt)"),
    expected_line_spacing: Optional[float] = typer.Option(None, "--expected-line-spacing", help="Expected line spacing"),
    expected_alignment: Optional[str] = typer.Option(None, "--expected-alignment", help="Expected alignment"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Verifies document integrity and typography conformance."""
    try:
        agent = DocumentAgent(file)
        res = agent.verify(
            expected_font=expected_font,
            expected_size_pt=expected_size,
            expected_line_spacing=expected_line_spacing,
            expected_alignment=expected_alignment,
        )
        output_result(res, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("diff")
def cmd_diff(
    before: str = typer.Argument(..., help="Path to before .docx file"),
    after: str = typer.Argument(..., help="Path to after .docx file"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Computes semantic diff between two DOCX revisions."""
    try:
        agent = DocumentAgent(before)
        diff_rep = agent.diff(after)
        output_result(diff_rep, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("apply-plan")
def cmd_apply_plan(
    plan_file: str = typer.Argument(..., help="Path to JSON plan file"),
    output: Optional[str] = typer.Option(None, "--output", help="Output file path"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Applies a multi-step batch operation plan atomically."""
    try:
        with open(plan_file, "r", encoding="utf-8") as f:
            plan_data = json.load(f)
        doc_path = plan_data.get("document")
        agent = DocumentAgent(doc_path)
        plan_res = agent.apply_plan(plan_data)
        out_p = agent.save(output_path=output)
        res = {"success": True, "operation": "apply_plan", "document": out_p, "details": plan_res}
        output_result(res, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("md2docx")
def cmd_md2docx(
    md_file: str = typer.Argument(..., help="Path to Markdown .md file"),
    output: Optional[str] = typer.Option(None, "--output", help="Output .docx file path"),
    preset: str = typer.Option("academic-vn", "--preset", help="Preset to style document"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Converts a Markdown file to a styled DOCX document."""
    try:
        out_docx = output or str(Path(md_file).with_suffix(".docx"))
        converter = MarkdownToDocxConverter(preset_name=preset)
        res_p = converter.convert(md_file, out_docx)
        res = {"success": True, "operation": "md2docx", "output": str(res_p)}
        output_result(res, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("workspace")
def cmd_workspace(
    file_path: Optional[str] = typer.Argument(None, help="Path to .docx file to open"),
    port: int = typer.Option(8765, "--port", "-p", help="HTTP port to host workspace"),
    no_browser: bool = typer.Option(False, "--no-browser", help="Do not automatically open browser"),
):
    """Launches the interactive AI-Native Document Workspace."""
    try:
        from docx_agent.interfaces.workspace.server import WorkspaceServer
        WorkspaceServer.launch(file_path=file_path, port=port, open_browser=not no_browser)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("visual-verify")
def cmd_visual_verify(
    file_path: str = typer.Argument(..., help="Path to .docx file"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Executes visual layout verification (overflows, whitespace, heading hierarchy)."""
    try:
        from docx_agent.adapters.docx import DocxImporter
        from docx_agent.verification.visual import VisualLayoutVerifier
        doc_node = DocxImporter.import_docx(file_path)
        rep = VisualLayoutVerifier.verify_document_layout(doc_node)
        output_result(rep, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("research")
def cmd_research(
    claim: str = typer.Argument(..., help="Factual or theoretical claim to research"),
    style: str = typer.Option("apa", "--style", help="Citation style (apa, ieee, academic-vn)"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Finds verified sources for a claim without hallucinating citations."""
    try:
        from docx_agent.research.provider import ResearchAssistant
        res_assistant = ResearchAssistant()
        proposal = res_assistant.evaluate_claim_and_propose_citation(claim, citation_style=style)
        output_result(proposal, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


@app.command("diagram")
def cmd_diagram(
    diagram_type: str = typer.Option("architecture", "--type", "-t", help="Diagram type (architecture, flowchart)"),
    title: str = typer.Option("System Architecture", "--title", help="Diagram title"),
    items: List[str] = typer.Option(..., "--item", "-i", help="Components or steps in diagram"),
    output_svg: Optional[str] = typer.Option(None, "--output-svg", help="Save rendered SVG to file"),
    as_json: bool = typer.Option(False, "--json", help="Output machine-readable JSON"),
):
    """Synthesizes structured Mermaid and SVG diagrams."""
    try:
        from docx_agent.media.diagrams import DiagramSynthesizer
        if diagram_type == "flowchart":
            diag = DiagramSynthesizer.generate_flowchart(items, title=title)
        else:
            diag = DiagramSynthesizer.generate_architecture_diagram(items, title=title)

        if output_svg and diag.rendered_svg:
            with open(output_svg, "w", encoding="utf-8") as f:
                f.write(diag.rendered_svg)

        output_result(diag, as_json=as_json)
    except Exception as e:
        err_console.print(f"[bold red]Error:[/bold red] {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    app()
