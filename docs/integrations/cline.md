# Cline & Roo Code & OpenAI Codex Integration Guide

## MCP Setup for Cline and Roo Code

Add to `cline_mcp_settings.json` or Roo Code settings:

```json
{
  "mcpServers": {
    "docx-agent": {
      "command": "docx-agent-mcp",
      "args": [],
      "disabled": false,
      "autoApprove": [
        "docx_inspect",
        "docx_read",
        "docx_outline",
        "docx_find",
        "docx_verify",
        "docx_diff"
      ]
    }
  }
}
```

## Python API Integration for OpenAI Codex & Custom Agents

```python
from docx_agent import DocumentAgent

agent = DocumentAgent("paper.docx")
agent.apply_preset("ieee")
agent.replace("legacy framework", "novel multi-agent framework")
agent.save("paper_updated.docx", verify=True)
```
