# Cursor IDE Integration Guide

## Overview

In Cursor, `docx-agent` can be configured as an MCP server or run from the integrated terminal.

## Cursor Settings (.cursor/mcp.json)

```json
{
  "mcpServers": {
    "docx-agent": {
      "command": "docx-agent-mcp"
    }
  }
}
```

## Agent Capabilities in Cursor

- Cursor AI can inspect Word outline hierarchies.
- Cursor AI can batch apply academic formatting presets.
- Cursor AI can execute text replacements with zero risk of breaking bold/italic/colored runs.
- Cursor AI can verify and diff document revisions directly.
