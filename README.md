# Email Organizer

Apply one controlled Gmail category label per message. Optionally archive older Inbox mail by removing the `INBOX` label. This project **never deletes email**.

## What it does

1. Reads Gmail messages (Trash, Spam, and Drafts are skipped).
2. Classifies each message with a local Ollama model.
3. Applies one label under `AI Mail/<Category>`.
4. Keeps messages up to 30 days old in Inbox.
5. Archives messages older than 30 days by removing `INBOX` only.

Default mode is a dry run: it shows what would happen and does **not** change Gmail.

## What you need

- Python 3.10+
- [Ollama](https://ollama.com/) running locally with `qwen2.5:3b` (or change `OLLAMA_MODEL` in `email_organizer.py`)
- A Google Cloud OAuth **Desktop** client with the `gmail.modify` scope

`gmail.modify` is required to create labels and to remove `INBOX`. It does not grant permanent delete.

## Setup

```text
pip install -r requirements.txt
```

1. In Google Cloud, create an OAuth client (Desktop app).
2. Download the client JSON.
3. Copy `credentials/credentials.json.example` to `credentials/credentials.json` and replace the placeholders with that client JSON.
4. Start Ollama and confirm the model is installed:

```text
ollama list
```

## Run

```text
python email_organizer.py
```

The first run opens a browser so you can authorize Gmail. The token is stored as `credentials/token.json` in this folder only.

Review the Excel file in `reports/` before you enable live changes.

### Enable live labeling / inbox archive

In `email_organizer.py`:

```python
DRY_RUN = False
```

Leave `DRY_RUN = True` until you have checked a report.

`PROCESS_LIMIT = 50` is a test cap. Set it to `None` to process all eligible mail.

## Safety

- Emails are never deleted.
- Trash, Spam, and Drafts are not processed.
- Older messages are only archived by removing `INBOX`.
- Category labels are applied under the `AI Mail` parent.

## Paths

All paths are relative to this project folder. There is no hardcoded machine-specific root.
