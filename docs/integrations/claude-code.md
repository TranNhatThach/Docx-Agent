# Claude Code Integration Guide

## Overview

Claude Code can utilize `docx-agent` through either its MCP tool server or direct CLI execution.

## MCP Server Setup

Add `docx-agent-mcp` to your Claude Desktop / Claude Code config:

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

## Recommended Workflow for Claude Code

1. **Inspect**: Call `docx_inspect(file="report.docx")` to understand outline and element count.
2. **Plan**: Propose operations (replacements, preset applications).
3. **Execute**: Call `docx_replace` or `docx_preset`.
4. **Verify**: Call `docx_verify` to confirm that typography, line spacing, and fonts match expected specifications.
