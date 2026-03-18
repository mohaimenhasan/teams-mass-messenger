# Teams Mass Messenger

A Windows desktop app (WPF) that sends a customizable message to multiple Microsoft Teams users via deep links and keyboard simulation. No Azure AD app registration required — it drives your local Teams desktop client directly.

## Prerequisites

- **Windows 10/11**
- **.NET 8 SDK** — [Download](https://dotnet.microsoft.com/download/dotnet/8.0)
- **Microsoft Teams desktop app** — must be open and signed in before running

## Project Structure

```
teams-mass-messenger/
├── README.md
├── alias.txt                    # List of Microsoft aliases (one per line)
├── message.txt                  # The message to send (edit this!)
├── MSSLTeamsMessenger.csproj
├── App.xaml / App.xaml.cs
├── MainWindow.xaml / .cs        # WPF UI and logic
├── Models/AliasEntry.cs         # Data model
├── Services/TeamsMessengerService.cs  # Teams automation
└── Helpers/AliasLoader.cs       # Alias file parser
```

## Setup

1. Clone the repo:
   ```
   git clone https://github.com/mohaimenhasan/teams-mass-messenger.git
   cd teams-mass-messenger
   ```

2. Restore and build:
   ```
   dotnet build
   ```

3. **Populate alias.txt** — this file ships empty (with a sample alias). Add one Microsoft alias per line for each person you want to message:
   ```
   johndoe
   janedoe
   someuser@microsoft.com
   ```
   **Important:** You must populate this file with the aliases you want to message before running the app.

   Rules:
   - One alias per line
   - Leading/trailing whitespace is trimmed automatically
   - Duplicates are removed (case-insensitive)
   - `@microsoft.com` suffix is optional (stripped during normalization)

4. Edit **message.txt** — write the message you want to send. Supports multiple lines:
   ```
   Hey! Want to join our team?
   We play on Wednesdays at 7 PM.
   ```

## Running

```
dotnet run
```

The app will open with:
- **Left panel** — Alias list loaded from `alias.txt` with checkboxes to select/deselect
- **Right panel** — Editable message (pre-loaded from `message.txt`), delay settings, and action buttons
- **Bottom** — Timestamped log of all actions

## Usage

### Test First
1. Open Microsoft Teams and make sure you're signed in
2. Click **"🧪 Test (Send to mohaimenkhan)"** to send the message to yourself
3. Verify the message arrived in your Teams chat
4. **Do not touch the computer** while sending — the app uses clipboard paste + keyboard simulation

### Send to Everyone
1. Select/deselect aliases using the checkboxes (or use Select All / Deselect All)
2. Click **"📨 Send to Selected"**
3. Confirm the dialog
4. **Do not use the computer** while batch sending — let it run
5. Use **"⛔ Stop"** to cancel mid-batch if needed

### Settings
- **Delay between messages** — Default is 7 seconds. Increase if Teams is slow to load chats. Decrease if you're feeling lucky.

## How It Works

1. For each alias, the app opens a Teams deep link: `https://teams.microsoft.com/l/chat/0/0?users=alias@microsoft.com`
2. Waits for Teams to open the 1:1 chat (configurable delay)
3. Focuses the Teams window using Win32 APIs
4. Pastes the message from clipboard (Ctrl+V) and presses Enter

## Troubleshooting

| Issue | Fix |
|-------|-----|
| Message lands in wrong window | Increase the delay; don't touch the computer during sending |
| Teams window not found | Make sure Teams desktop app is open (not just the browser) |
| alias.txt not found | Place it in the project root folder |
| Message is empty | Check that message.txt exists and has content |

## License

MIT
