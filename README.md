# create_wsone_api_bruno_collection

PowerShell utility to generate a Bruno API collection from a Workspace ONE UEM server.

## What It Does

The script:

1. Prompts for (or accepts) Workspace ONE server and OAuth credentials.
2. Requests an OAuth2 access token.
3. Reads server version from `/api/system/info`.
4. Creates a versioned Bruno collection folder.
5. Downloads OpenAPI docs and converts endpoints into `.bru` requests.
6. Configures generated requests to inherit auth.

## Requirements

- PowerShell 5.1 or later
- Workspace ONE UEM with OAuth2 client credentials
- Access to the UEM API endpoints

## Script

- `create-bruno copy.ps1`

## Usage

Run with prompts:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\create-bruno copy.ps1
```

Run with parameters:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\create-bruno copy.ps1 `
  -clientId "YOUR_CLIENT_ID" `
  -clientSecret "YOUR_CLIENT_SECRET" `
  -tokenUrl "https://your-oauth-host/connect/token" `
  -Server "https://apiXXX.awmdm.com"
```

## Output

The script generates Bruno collection folders and `.bru` request files based on the server's API docs.

## License

This project is licensed under the MIT License. See `LICENSE`.