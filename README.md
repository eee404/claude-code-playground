# Claude Code Playground

A Windows launcher for Claude Code conversations in VS Code. Double-click, enter a topic, and VS Code opens with Claude Code ready in the terminal.

## Installation

1. Clone the repo into a fixed folder:
   ```
   git clone <url> C:\Tools\claude-code-playground
   ```

2. (Optional) Copy `.env.template` to `.env` to customize the conversations folder:
   ```
   CONVERSATIONS_DIR=C:\path\to\my\conversations
   ```
   By default, conversations are created in a `conversations/` folder next to the script.

3. (Optional) Create a shortcut to `new-conversation.vbs` on the Desktop for quick access.

## Usage

1. Double-click `new-conversation.vbs`
2. Enter a topic (or leave blank)
3. VS Code opens in the new folder with Claude Code launched in the terminal

Conversations are created as `YYYY-MM-DD_NNN_slugified-topic/` (e.g. `2026-03-23_001_my-topic`, `2026-03-23_002`). The index auto-increments per day.

## Customization

- **`CLAUDE.md.template`**: the system prompt copied into each new conversation. Edit it to suit your needs. See `CLAUDE.md.template.example` for an example.
- **`.env`**: path to the conversations folder. Accepts an absolute path (`C:\...`) or a relative path (resolved relative to the script).

## Recommended VS Code Configuration

For Claude Code to launch automatically without confirmation on each new conversation:

1. **Allow automatic tasks**: `Ctrl+,` in VS Code → search for "automatic tasks" → set to **On**

2. **Trust the conversations folder**: open any conversation in VS Code, then `Ctrl+Shift+P` → "Workspaces: Manage Workspace Trust" → add the parent `conversations/` folder as trusted. All subfolders will automatically be trusted.

## Conversation Structure

```
conversations/
└── 2026-03-23_001_my-topic/
    ├── CLAUDE.md          (copied from the template)
    └── .vscode/
        └── tasks.json     (auto-launches claude in the terminal)
```
