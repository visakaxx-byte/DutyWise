# BParty V2 Server Deploy

Target directory:

```text
/opt/bparty-v2
```

Runtime:

```bash
python3 -m venv /opt/bparty-v2/.venv
/opt/bparty-v2/.venv/bin/pip install -r /opt/bparty-v2/requirements.txt
cp /opt/bparty-v2/deploy/bparty-v2.service /etc/systemd/system/bparty-v2.service
systemctl daemon-reload
systemctl enable --now bparty-v2
```

Required `/opt/bparty-v2/.env`:

```text
CRAWLER_USERNAME=...
CRAWLER_PASSWORD=...
CRAWLER_BASE_URL=https://www.codeflagai.com
```

`LLM_*` / `DOUBAO_*` can stay configured for future extensions, but the current target-tax flow only requires codeflagai credentials.
