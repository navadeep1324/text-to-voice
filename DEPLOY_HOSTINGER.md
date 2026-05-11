# Hostinger Deployment Checklist

## Recommended n8n schedule

Use n8n as the scheduler. Keep the deployed app as an on-demand API worker only.

In n8n:

1. Add a Schedule Trigger.
2. Set timezone to `Asia/Kolkata`.
3. Run at `7:30 PM`.
4. Run only Monday, Tuesday, Wednesday, Thursday, and Friday.
5. Add an HTTP Request node:

```text
POST https://your-domain.com/api/n8n/generate-audio
```

Headers:

```text
Authorization: Bearer your-random-n8n-api-key
```

The API response contains generated MP3 files as base64 in:

```text
results[].base64
```

Each item also includes:

```text
results[].filename
results[].contentType
results[].sheet
```

No MP3 file is saved on Hostinger in this API mode.

Use n8n to convert each `base64` value into binary data and send/upload it to Microsoft Teams.

## Optional Hostinger Cron schedule

Run the one-shot cron script from Hostinger Cron instead of enabling the internal Next.js scheduler.

7:30 PM IST is 2:00 PM UTC, and Hostinger Cron uses UTC.

```cron
0 14 * * 1-5
```

Cron command:

```sh
/bin/sh /home/YOUR_USER/domains/YOUR_DOMAIN/public_html/scripts/hostinger-cron.sh
```

## Required environment variables

```env
N8N_API_KEY=your-random-n8n-api-key
SP_TENANT_ID=your-tenant-id
SP_CLIENT_ID=your-client-id
SP_CLIENT_SECRET=your-sharepoint-client-secret
SP_HOST=yourtenant.sharepoint.com
SP_SITE_PATH=/sites/YourSite
SP_FILE_PATH=Shared Documents/path/to/workbook.xlsx
SP_SHEETS=AI Team,Dev Team,DM Team
```

`TEAMS_WEBHOOK_URL` is only needed if you use the optional direct cron/webhook mode. The n8n API mode does not need it.

If `config.json` is present in the app working directory, the script reads non-secret SharePoint settings from it. On Hostinger, the Node.js app may run from a `nodejs` working directory where `config.json` is not present, so environment variables are the safest deployment option.

## Config for link-based webhook mode only

Set the public URL used in Teams messages:

```json
"teams": {
  "public_base_url": "https://your-domain.com"
}
```

## SharePoint permissions

The Microsoft Entra app must be able to read and write the configured workbook.

Required Graph application permission:

```text
Sites.ReadWrite.All
```

Then grant admin consent.

`Sites.Read.All` is enough only for downloading the workbook.

## Runtime behavior

When n8n calls the API, the app:

1. Downloads the SharePoint Excel workbook.
2. Checks row 2 in each configured sheet.
3. Generates audio only if row 2 has today's IST date.
4. Skips sheets that do not have today's date or do not produce voice text.
5. Returns generated MP3 files to n8n as base64.
6. Does not save MP3 files to Hostinger storage.

Incoming Teams webhooks can send JSON card messages, but they cannot upload MP3 binary files. For actual files in Teams, let n8n use a Microsoft Teams or Microsoft Graph connection to upload/send the binary files it receives from the API.

Run a manual test before enabling Cron:

```sh
node scripts/generate-sharepoint-audio.cjs
```
