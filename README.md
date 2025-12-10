# OMSendTest - Outlook OnMessageSend Add-in

An Outlook add-in that uses the OnMessageSend event to check for missing attachments before sending emails.

## Features

- Detects keywords like "attachment", "document", "picture", or "send" in email body
- Prompts user if these keywords are found but no attachments are present
- Prevents accidental sending of emails without intended attachments
- Works with Outlook Desktop (Windows)

## Prerequisites

- Node.js (v14 or later)
- npm
- Outlook Desktop (Classic Outlook on Windows)
- Microsoft 365 account for testing

## Installation

1. Clone the repository:
```bash
git clone https://github.com/SenMS-CSA/OMSendTest.git
cd OMSendTest
```

2. Install dependencies:
```bash
npm install
```

3. Trust the development certificates:
```bash
npx office-addin-dev-certs install
```

## Development

### Build the project

Development build:
```bash
npm run build:dev
```

Production build:
```bash
npm run build
```

### Run the add-in

Start the dev server and sideload the add-in in Outlook Desktop:
```bash
npm start -- desktop --app outlook
```

Or use the default configured app:
```bash
npm start
```

### Stop debugging

```bash
npm stop
```

### Validate the manifest

```bash
npm run validate
```

## Project Structure

```
OMSendTest/
├── assets/              # Icons and images
├── src/
│   ├── commands/        # Command functions
│   ├── launchevent/     # OnMessageSend event handler
│   └── taskpane/        # Task pane UI
├── manifest.xml         # Add-in manifest
├── webpack.config.js    # Webpack configuration
└── package.json         # Project dependencies
```

## How It Works

The add-in uses the OnMessageSend event to intercept email sends:

1. When user clicks Send, the `onMessageSendHandler` function is triggered
2. The email body is checked for keywords: "send", "picture", "document", "attachment"
3. If keywords are found, it checks for attachments
4. If no attachments are present, it blocks the send and prompts the user
5. User can choose to send anyway or cancel

## Important Implementation Notes

### Webpack Configuration

The event handler JavaScript file (`launchevent.js`) is copied directly to avoid webpack dev server's hot module replacement code, which interferes with Office.js runtime:

```javascript
// In webpack.config.js
{
  from: "src/launchevent/launchevent.js",
  to: "launchevent-src.js",
}
```

### Manifest Configuration

The manifest points to the copied source file for classic Outlook:

```xml
<bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/launchevent-src.js" />
```

## Troubleshooting

### Event not triggering

- Ensure the dev server is running on port 3000
- Check runtime logs: `%LOCALAPPDATA%\Temp\OfficeAddins.log.txt`
- Verify the manifest is properly loaded in Outlook
- Close and reopen compose windows after changes

### "Taking longer than expected" error

This usually indicates the event handler isn't completing. Check:
- `Office.actions.associate()` is being called
- Event handler calls `event.completed()` in all code paths
- No JavaScript errors in the runtime

## License

MIT

## Support

For issues and questions, please open an issue in the GitHub repository.
