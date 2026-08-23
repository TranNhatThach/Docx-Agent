# Antigravity Integration Guide

## Overview

Antigravity seamlessly integrates with `docx-agent` through both the bundled MCP server and CLI tool adapters.

## MCP Configuration

Add the following configuration to your `mcp_config.json`:

```json
{
  "mcpServers": {
    "docx-agent": {
      "command": "docx-agent-mcp",
      "args": []
    }
  }
}
```

## Available Tools

Once connected, Antigravity has access to:
- `docx_inspect`: Read document summaries and heading hierarchies.
- `docx_read`: Inspect paragraphs by stable ID (`p_0001`, `p_0002`...).
- `docx_find`: Search text, regex, or styles.
- `docx_replace`: Surgical format-preserving replacements.
- `docx_preset`: One-click institutional formatting (`academic-vn`, `ieee`, `apa`).
- `docx_verify`: Independent verification of document integrity and typography.
- `docx_diff`: Structural revision comparison.

## CLI Usage

Alternatively, Antigravity can execute commands directly:

```bash
docx-agent inspect document.docx --json
docx-agent replace document.docx --target "old" --replace "new" --json
docx-agent verify document.docx --expected-font "Times New Roman" --json
```
