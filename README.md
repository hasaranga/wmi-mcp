# WMI MCP Server

A Model Context Protocol (MCP) server that enables LLMs to execute Windows Management Instrumentation (WMI) queries to retrieve comprehensive system information from Windows computers.

## Overview

This is a **stdio-based MCP server** that provides LLMs with the ability to query Windows systems for:
- Hardware information (CPU, memory, disks, network adapters)
- Running processes and services
- System configuration and settings
- Network information
- Event logs
- Performance counters
- And much more through WMI

## Features

- Execute WQL (WMI Query Language) queries against any WMI namespace
- Returns structured JSON data for easy LLM parsing
- Supports all major WMI namespaces (ROOT\CIMV2, ROOT\WMI, etc.)
- Comprehensive error handling
- Type-safe data conversion for all WMI property types

## Installation

1. **Download from releases or compile the executable** (`wmi-mcp.exe`)
2. **Place the executable** in a directory of your choice (e.g., `C:\Users\YourName\wmi-mcp\`)
3. **Configure your LLM agent** to use this MCP server

### Compilation (if building from source)

```bash
# Using MSVC
cl main.cpp /EHsc ole32.lib oleaut32.lib wbemuuid.lib

# Using MinGW/GCC
g++ -std=c++11 main.cpp -lole32 -loleaut32 -lwbemuuid -o wmi-mcp.exe
```

## Configuration

### For Gemini CLI

Add the following to your MCP configuration:

```json
{
  "mcpServers": {
    "WMIServer": {
      "command": "wmi-mcp.exe",
      "cwd": "C:\\Users\\XXX\\wmi-mcp",
      "timeout": 15000,
      "trust": true
    }
  }
}
```

### For Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "wmi-query": {
      "command": "C:\\Users\\XXX\\wmi-mcp\\wmi-mcp.exe",
      "args": []
    }
  }
}
```

### For Other MCP Clients

Configure as a stdio server pointing to the executable path.

## Ask Questions About Your System

Once set up, you can simply ask the LLM questions like "Why is my computer running hot?" or "Is my hard drive failing?". The LLM will translate your question into the necessary WMI queries and send them to this MCP server for execution.

### Few Examples

```
Why does my PC boot slow?
Why is my computer running hot?
Why is my internet slow?
Why is my battery draining so fast?
Why is my game lagging?
Were there failed or unexpected login attempts?
Is my hard drive failing?
Is there any suspicious process on my PC?
Can my PC run this game/software?
Which programs start with Windows?
What processes are using my network right now?
Is my antivirus active and up to date?
When was the last time my system updated?
Are critical Windows services disabled or tampered with?
Has new software installed recently?
Are there suspicious scheduled tasks?
Are my firewall and antivirus still active?
Which programs connected to suspicious IP addresses?
Are there suspicious startup programs?
```

## Tool Parameters

### executeWMIQuery

**Parameters:**
- `namespace` (optional): WMI namespace to query
  - Default: `ROOT\CIMV2`
  - Examples: `ROOT\WMI`, `ROOT\RSOP\Computer`, `ROOT\SecurityCenter2`
- `query` (required): WQL query string
  - Example: `SELECT * FROM Win32_Process`

**Returns:**
Structured JSON with:
- `success`: Boolean indicating query success
- `namespace`: The namespace that was queried
- `query`: The executed query
- `count`: Number of objects returned
- `objects`: Array of WMI objects with their properties
- `error`: Error message (if any)

## Common WMI Classes

| Class | Description |
|-------|-------------|
| `Win32_Process` | Running processes |
| `Win32_Service` | System services |
| `Win32_LogicalDisk` | Disk drives |
| `Win32_ComputerSystem` | Computer system info |
| `Win32_OperatingSystem` | OS information |
| `Win32_Processor` | CPU information |
| `Win32_PhysicalMemory` | RAM modules |
| `Win32_NetworkAdapter` | Network adapters |
| `Win32_NetworkAdapterConfiguration` | Network configuration |
| `Win32_EventLog` | Event log files |
| `Win32_NTLogEvent` | Event log entries |

## WMI Namespaces

| Namespace | Purpose |
|-----------|---------|
| `ROOT\CIMV2` | Most common system information |
| `ROOT\WMI` | Hardware sensors and performance |
| `ROOT\SecurityCenter2` | Security products info |
| `ROOT\RSOP` | Group Policy information |
| `ROOT\Microsoft\Windows\Storage` | Storage management |

## Security Notes

- The server will not start if executed with administrative privileges, as a safety measure.
- Consider the security implications of allowing LLMs to query system information

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve this MCP server.

## Author

R.Hasaranga (https://www.hasaranga.com)

## Privacy Policy

This program will not transfer any information to other networked systems unless specifically requested by the user or the person installing or operating it.

## Code Signing Policy

Free code signing provided by SignPath.io, certificate by SignPath Foundation.


