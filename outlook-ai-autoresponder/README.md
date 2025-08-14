### Outlook AI Autoresponder (Microsoft 365 / Outlook via Microsoft Graph)

A small Python service that auto-replies to new Outlook emails with professional, context-aware AI responses. It uses Microsoft Graph for mail access and OpenAI or Azure OpenAI to generate replies. Runs on Linux as a background process and works even if you never open the email.

- **Automatic replies**: Responds to new incoming emails without user interaction
- **AI-generated**: Professional replies tailored to the sender and content
- **Loop-safe**: Skips auto-generated and no-reply senders; tags processed messages
- **Privacy**: Processes only your mailbox
- **No public webhook required**: Uses polling with Device Code auth

---

### Prerequisites

- A Microsoft 365 mailbox (Work/School or Personal with Outlook.com)
- Azure Entra ID App Registration (no server needed; Device Code flow)
- Python 3.10+
- OpenAI API key or Azure OpenAI resource

---

### Microsoft Entra ID App Registration

1) Go to `https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/overview`
2) Azure Active Directory → App registrations → New registration
   - Name: `Outlook AI Autoresponder`
   - Supported account types: `Accounts in this organizational directory only` (or as needed)
   - Redirect URI: Not needed for Device Code
3) After creating, copy:
   - Application (client) ID → `AZURE_CLIENT_ID`
   - Directory (tenant) ID → `AZURE_TENANT_ID`
4) Authentication → Advanced settings → Allow public client flows → Enable (Yes)
5) API permissions → Add a permission → Microsoft Graph → Delegated permissions:
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `offline_access`
   - (Optional) `Contacts.Read`
6) Grant admin consent if your tenant requires it for these delegated scopes.

---

### Configure AI Provider

Use one of the following:

- OpenAI
  - Set `OPENAI_API_KEY`
  - Set `LLM_MODEL` (e.g., `gpt-4o-mini`)

- Azure OpenAI
  - Set `AZURE_OPENAI_API_KEY`
  - Set `AZURE_OPENAI_ENDPOINT` (e.g., `https://YOUR-RESOURCE.openai.azure.com`)
  - Set `AZURE_OPENAI_API_VERSION` (e.g., `2024-06-01`) 
  - Set `LLM_MODEL` to your deployment name (the model deployment ID in Azure OpenAI)

---

### Setup

```bash
cd /workspace/outlook-ai-autoresponder
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
# Edit .env and fill in IDs and API keys
```

Run once to authenticate (Device Code prompts in the terminal):

```bash
python app.py --once
```

Then run continuously:

```bash
python app.py
```

Optional dry run (no emails are sent, just logs):

```bash
python app.py --dry-run
```

---

### Environment Variables (`.env`)

- `AZURE_TENANT_ID`: Entra tenant ID
- `AZURE_CLIENT_ID`: App registration client ID
- `OPENAI_API_KEY`: OpenAI key (if using OpenAI)
- `AZURE_OPENAI_API_KEY`: Azure OpenAI key (if using Azure OpenAI)
- `AZURE_OPENAI_ENDPOINT`: Azure OpenAI endpoint URL
- `AZURE_OPENAI_API_VERSION`: Azure OpenAI API version
- `LLM_MODEL`: Model or deployment name (e.g., `gpt-4o-mini` or your Azure deployment)
- `POLL_INTERVAL_SECONDS`: Polling interval (default: 30)
- `STATE_FILE`: Persistent state file path (default: `state.json`)
- `TOKEN_CACHE_FILE`: MSAL token cache file path for silent re-auth (default: `token_cache.json`)
- `ALLOWED_LANGUAGES`: Optional comma-separated ISO codes (e.g., `en,fr,de`) to constrain reply language
- `USER_NAME`: Your name for signature/context (optional)
- `USER_ROLE`: Your role/title (optional)
- `ORG_NAME`: Organization name (optional)
- `REPLY_SIGNATURE`: Custom HTML signature appended to replies (optional)

---

### What it does

- Polls your Inbox for newly received messages since the last check
- Skips messages likely to be automated (subjects like "Automatic reply", `no-reply@` senders)
- Generates a concise, professional response using LLM
- Replies to the sender in-thread
- Tags the original message with category `AutoReplied` (created if missing)
- Keeps a local `state.json` to avoid duplicate replies

---

### Safety and Loop Avoidance

- Skips if the sender appears to be automated (e.g., `no-reply`) or the subject indicates OOF/autoreply
- Skips messages from yourself
- Adds `AutoReplied` category to processed items
- Maintains a recent set of `replied_ids` to avoid duplicate sends

---

### Running as a service (systemd)

Example unit file `/etc/systemd/system/outlook-ai-autoresponder.service`:

```ini
[Unit]
Description=Outlook AI Autoresponder
After=network-online.target

[Service]
Type=simple
WorkingDirectory=/workspace/outlook-ai-autoresponder
Environment=PYTHONUNBUFFERED=1
ExecStart=/workspace/outlook-ai-autoresponder/.venv/bin/python /workspace/outlook-ai-autoresponder/app.py
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now outlook-ai-autoresponder
journalctl -u outlook-ai-autoresponder -f
```

---

### Notes

- First run prompts a device code flow; subsequent runs are silent using the token cache file
- You can adjust content/style by changing the prompt in `app.py`
- To stop auto-replying temporarily, stop the service or run with `--dry-run`

---

### Troubleshooting

- If authentication fails, ensure public client flows are enabled and delegated scopes are consented
- For Azure OpenAI, ensure the `LLM_MODEL` matches your deployment name and API version is correct
- If replies don’t send, verify `Mail.Send` and `Mail.ReadWrite` delegated permissions are granted
