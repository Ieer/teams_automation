# Teams Automation

Helpers for automating the Microsoft Teams desktop client on Windows using `uiautomation`.

## Installation

```powershell
pip install -e .
```

## Quick start

```python
from teams_automation import TeamsClient

client = TeamsClient(minimize_after_send=True)

# Send a simple text message
client.send_text(
    "Hello from automation!",
    chat_name="Columbus Teams Chatbot",
)

# Send one or more files, optionally with a caption
client.send_files(
    ["team_tree.png"],
    chat_name="Columbus Teams Chatbot",
    caption="Team org chart",
)
```

## Parameters

- `TeamsClient`
  - `window_keyword`: keywords used to locate the Teams window title.
  - `window_class_keyword`: control class name filter; keep default for Desktop client.
  - `activation_delay`: seconds to wait after activating the window before interacting.
  - `search_timeout`: time to wait after entering a chat filter before inspecting results.
  - `section_preference`: priority order of sidebar sections (for example `"Chats"`, `"Favorites"`).
  - `aliases`: mapping to normalise chat names returned by Teams, useful when names vary.
  - `minimize_after_send`: optionally minimises windows once the message or files are sent.

- `send_text` / `send_message`
  - `message`: text content to post.
  - `chat_name`: target chat display name (normalised with `aliases`).
  - `section_name`: restrict search to a specific sidebar section; defaults to preferred order.
  - `image_path` (`send_message` only): optional PNG/JPG/etc. inserted inline with the text.
  - `close_filter`: close the Teams search box after selecting the chat.
  - `wait_after_send`: pause after clicking send to let the UI settle.

- `send_files`
  - `filepaths`: iterable of files to upload.
  - `caption`: optional text typed into the message box before attachments are pasted.
  - `embed_images`: paste image files inline (default `True`); when `False` they upload as attachments.
  - `chat_name`, `section_name`, `close_filter`, `wait_after_send`: behave the same as in `send_message`.

## Development

- Requires Windows with Microsoft Teams desktop client.
- Python 3.9 or newer.
- Install dev dependencies: `pip install -e .`

## Caveats

- UI element names may vary based on Teams build and language settings.
- Automation relies on focus; avoid interacting with the machine during runs.
