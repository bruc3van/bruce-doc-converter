# Install and update

Check installation and updates once per session. Follow the host's install approval, audit and network rules, and any version, location or offline requirement the user gave. When the host requires review before installing, finish that review first; do not ask again when existing approval already covers it.

## 1. Check the current installation

```bash
bdc --help-json
```

- Success returns JSON with `cli_version` and, from 0.2.2, `install`:

  | `install.source` | Meaning | `auto_upgrade` |
  | --- | --- | --- |
  | `pipx` | pipx-managed tool environment | `true` |
  | `uv` | `uv tool` environment | `true` |
  | `user` | `pip install --user` | `true` |
  | `venv` | a virtual environment, possibly a project dependency | `false` |
  | `system` | system Python site-packages | `false` |
  | `source` | a development checkout run from source | `false` |

  `upgrade_command` is the matching upgrade command, or `null` when the CLI should not upgrade itself.

- `bdc --version` prints only the version (0.2.2 and later).
- Command not found: check `.venv/bin/bdc` (macOS/Linux) or `.venv\Scripts\bdc` (Windows) in the project before concluding it is missing. Being absent from PATH is not the same as not installed.
- CLIs before 0.2.2 have no `install` field (and no `--version`). Identify the installation from `pipx list --short`, `uv tool list`, or the path of the `bdc` command (`command -v bdc`, or `where bdc` on Windows): a path inside `.venv` is a project environment; otherwise treat it as `pip --user` only if it is under the user's script directory, and as `system` if unsure.

## 2. Find the latest release

```bash
python -m pip index versions bruce-doc-converter
```

The first line is `bruce-doc-converter (<latest>)`. This respects the user's configured package index and mirrors. If `pip index` is unavailable in an old pip, query PyPI directly:

```bash
python -c "import json,urllib.request;print(json.load(urllib.request.urlopen('https://pypi.org/pypi/bruce-doc-converter/json',timeout=10))['info']['version'])"
```

Use `python3` where `python` is not available. Compare semantic versions (`0.10.0` is newer than `0.9.0`). If the installed version is equal or newer, use it as is; never downgrade or pick a pre-release.

## 3. Decide whether to update

| Installation | Newer release found |
| --- | --- |
| `auto_upgrade: true` | Run `upgrade_command`, re-run `bdc --help-json`, and tell the user the old and new versions. |
| `venv` | Tell the user an update exists. Run `upgrade_command` only after they agree; do not change project lock files on your own. |
| `system` / `source` | Tell the user; do not modify the installation. |
| User- or host-pinned version | Keep it; only mention the update. |
| Not installed | Install the latest release (section 4). |

A newer release is not automatically approved by the host. When the host requires audits, a minimum release age or approvals, follow that process; if it cannot be satisfied, keep the installed version. If the query fails or you are offline, keep using an installed CLI and say updates were not checked; if nothing is installed, report the actual blocker.

Manual upgrade commands, matching how the CLI was installed:

```bash
pipx upgrade bruce-doc-converter
uv tool upgrade bruce-doc-converter
python -m pip install --user --upgrade bruce-doc-converter
.venv/bin/pip install --upgrade bruce-doc-converter        # macOS/Linux venv
.venv\Scripts\pip install --upgrade bruce-doc-converter    # Windows venv
```

After upgrading, run `bdc setup-node` if Markdown to Word is used; its Node.js dependencies may have changed. Conversions also report this as `DEPENDENCY_INSTALL_REQUIRED`.

## 4. Install

Try in order and stop at the first that succeeds:

```bash
pipx install bruce-doc-converter          # preferred: isolated, bdc on PATH
uv tool install bruce-doc-converter       # isolated, bdc on PATH
python -m pip install --user bruce-doc-converter
python -m venv .venv && .venv/bin/pip install bruce-doc-converter   # Windows: .venv\Scripts\pip
```

With the venv fallback, `bdc` is not on PATH: use `.venv/bin/bdc` (macOS/Linux) or `.venv\Scripts\bdc` (Windows) for every command. Reuse the same entry point for the rest of the session.

## 5. Markdown to Word runtime

Markdown to Word requires **Node.js >=22.0**; Office/PDF to Markdown does not use Node.js. If Node.js is missing or too old, tell the user what is needed; reinstalling dependencies cannot fix an old runtime, and do not switch the system's default Node.js yourself.

```bash
bdc setup-node
```

This installs the locked Node.js dependencies into a shared user directory with `npm ci --ignore-scripts`. It is idempotent and returns `already_installed: true` when nothing changed.

Mermaid diagrams render through the local Chrome, Edge or Chromium, launched headless with a temporary profile. Set `BRUCE_DOC_CONVERTER_CHROME_PATH` to choose a browser. Only if no local browser exists, download Puppeteer's browser:

```bash
bdc setup-node --install-browser
bdc setup-node --allow-scripts --install-browser   # only if the environment requires npm lifecycle scripts
```

On Linux sandboxes where Chromium cannot start, `BRUCE_DOC_CONVERTER_ALLOW_CHROMIUM_NO_SANDBOX=1` disables the Chromium sandbox. Set it only when the user accepts that risk.

## Troubleshooting

| Error | Cause | What to do |
| --- | --- | --- |
| `SOCKS support` / proxy connection error | `all_proxy` / `http_proxy` point to a proxy pip cannot use | Tell the user. With their agreement, clear the proxy variables for that one command only; do not change persistent proxy settings. |
| `command not found: pipx` | pipx not installed | Try `uv tool install` or `pip install --user`. |
| `externally-managed-environment` | System Python forbids pip installs | Use pipx, `uv tool install` or the venv fallback; do not pass `--break-system-packages`. |
| Permission denied | No write access to the install location | Use `--user`, pipx, uv or a venv; do not use `sudo` without the user's agreement. |
| `bdc: command not found` after a venv install | venv scripts are not on PATH | Use the full path `.venv/bin/bdc` or `.venv\Scripts\bdc`. |
