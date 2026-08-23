"""
Diagram Synthesizer & SVG Renderer: Generates system architectures, flowcharts,
sequence diagrams, and ER diagrams.
"""

from typing import Dict, Any, Optional, List
from pathlib import Path
from pydantic import BaseModel, Field

from docx_agent.canonical.model import DiagramBlock, ProvenanceRecord, ProvenanceType, generate_id


class DiagramSynthesizer:
    """
    Synthesizes structured Mermaid source definitions and clean SVG vectors.
    """

    @staticmethod
    def generate_architecture_diagram(
        components: List[str],
        title: str = "System Architecture",
        direction: str = "TB",
    ) -> DiagramBlock:
        """Generates a multi-tier system architecture diagram."""
        lines = [f"graph {direction}", f"    subgraph {title}"]
        for idx, comp in enumerate(components):
            safe_id = f"C{idx+1}"
            lines.append(f'        {safe_id}["{comp}"]')
            if idx > 0:
                prev_id = f"C{idx}"
                lines.append(f"        {prev_id} --> {safe_id}")
        lines.append("    end")
        mermaid_code = "\n".join(lines)

        # Generate clean SVG
        svg_content = f"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 600 {100 + len(components)*50}" width="100%" height="100%">
  <rect width="100%" height="100%" fill="#F8FAFC" rx="8" stroke="#CBD5E1" stroke-width="2"/>
  <text x="20" y="35" font-family="Arial, sans-serif" font-size="16" font-weight="bold" fill="#0F172A">{title}</text>
"""
        y = 60
        for idx, comp in enumerate(components):
            svg_content += f"""  <rect x="50" y="{y}" width="500" height="40" fill="#E2E8F0" rx="6" stroke="#94A3B8"/>
  <text x="300" y="{y+25}" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#1E293B">{comp}</text>
"""
            if idx < len(components) - 1:
                svg_content += f"""  <line x1="300" y1="{y+40}" x2="300" y2="{y+50}" stroke="#64748B" stroke-width="2" marker-end="url(#arrow)"/>
"""
            y += 50

        svg_content += "</svg>"

        return DiagramBlock(
            diagram_type="architecture",
            source_code=mermaid_code,
            rendered_svg=svg_content,
            caption=f"Hình: {title}",
            provenance=ProvenanceRecord(
                source_type=ProvenanceType.AGENT,
                creator="diagram_agent",
                notes=f"Generated architecture diagram with {len(components)} layers.",
            ),
        )

    @staticmethod
    def generate_flowchart(steps: List[str], title: str = "Process Flow") -> DiagramBlock:
        """Generates a sequential flowchart."""
        lines = ["graph LR"]
        for idx, step in enumerate(steps):
            safe_id = f"S{idx+1}"
            lines.append(f'    {safe_id}["{step}"]')
            if idx > 0:
                lines.append(f"    S{idx} --> S{idx+1}")
        mermaid_code = "\n".join(lines)

        svg_content = f"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {max(600, len(steps)*160)} 120" width="100%" height="100%">
  <rect width="100%" height="100%" fill="#F8FAFC" rx="8" stroke="#E2E8F0"/>
"""
        x = 20
        for idx, step in enumerate(steps):
            svg_content += f"""  <rect x="{x}" y="35" width="120" height="50" fill="#3B82F6" rx="6"/>
  <text x="{x+60}" y="65" font-family="Arial, sans-serif" font-size="12" fill="#FFFFFF" text-anchor="middle">{step}</text>
"""
            if idx < len(steps) - 1:
                svg_content += f"""  <line x1="{x+120}" y1="60" x2="{x+150}" y2="60" stroke="#94A3B8" stroke-width="2"/>
"""
            x += 150
        svg_content += "</svg>"

        return DiagramBlock(
            diagram_type="flowchart",
            source_code=mermaid_code,
            rendered_svg=svg_content,
            caption=f"Hình: {title}",
            provenance=ProvenanceRecord(
                source_type=ProvenanceType.AGENT,
                creator="diagram_agent",
            ),
        )
