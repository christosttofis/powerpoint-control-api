# PowerPoint Control API

A tiny HTTP server that lets other devices on the same network control **Microsoft PowerPoint
on a Windows laptop**: open a deck, start the slideshow, go to next / previous / any slide,
and read the current slide number.

Anything that can send an HTTP request can drive the presentation: a phone, another PC,
a script, `curl`, or your own app.

---

## What you need

| Requirement | Notes |
|---|---|
| **Windows 10 or 11** | The server talks to PowerPoint through Windows COM, so it only runs on Windows. |
| **Microsoft PowerPoint (desktop app)** | Microsoft 365 or Office 2016+. PowerPoint Online (the web version) does **not** work. |
| **Internet (first setup only)** | `setup.bat` downloads [uv](https://docs.astral.sh/uv/), Python and the dependencies. You do **not** need to install Python yourself. |
| **Same network** | The laptop and the controlling device must be on the same Wi-Fi / LAN. |

Everything is installed automatically with **uv**: its own Python 3.12 plus `flask`
(web server) and `pywin32` (the "win32" library that lets Python control PowerPoint),
in a local `.venv` folder. Nothing is installed system-wide except the small `uv` tool.

---

## Setup in 5 minutes

1. **Download** this repo: green **Code** button → **Download ZIP**, then right-click the
   ZIP → **Extract All…** (e.g. to your Desktop).
   *(Or with git: `git clone https://github.com/christosttofis/powerpoint-control-api.git`)*

2. **Double-click `setup.bat`** (one-time only).
   - If Windows shows *"Windows protected your PC"* → click **More info** → **Run anyway**.
   - It installs uv, Python and the dependencies (about 1 minute on first run).
   - When asked *"Add a Windows Firewall rule now?"* press **Y** and accept the admin prompt.
     This lets other devices reach the laptop on port 8000.

3. **Double-click `start.bat`** every time you present. A black window opens and shows
   something like:

   ```
   [PowerPoint Control API] Starting on port 8000
   [PowerPoint Control API] Clients should connect to one of:
       192.168.1.50:8000
   ```

   **Keep this window open** during the presentation. Close it (or press Ctrl+C) to stop.

That's it. The laptop is ready to be controlled.

---

## Presenting

1. Connect the laptop and the controlling device to the **same Wi-Fi**.
2. Run `start.bat` and note the address it prints (e.g. `192.168.1.50:8000`).
3. Enter that address as the PowerPoint host in whatever app/script controls the slides.
4. Get the deck on screen, either:
   - upload it through `/open` (the laptop opens it and starts the slideshow automatically), **or**
   - open the `.pptx` on the laptop yourself and press **F5** (the slideshow must be running).
5. Change slides from the controlling device.

**Tips**
- Turn off sleep/screen-off on the laptop (Settings → System → Power).
- Plug the laptop into the projector **before** starting the slideshow.
- If your controller uses slide numbers, make sure it's the same `.pptx` it was set up for.

---

## API reference

All responses are JSON. Default port: `8000`.

| Method | Path | What it does |
|---|---|---|
| `GET` | `/ping` | Health check. Returns current slide and total slides if a deck is open. Does not launch PowerPoint. |
| `POST` | `/next` | Next slide (or next animation step). |
| `POST` | `/prev` | Previous slide. |
| `POST` | `/goto/<n>` | Jump to slide `n` (1-based). |
| `POST` | `/open` | Upload a `.pptx` (multipart field `file`), close other decks, open it and start the slideshow. |

`/next`, `/prev` and `/goto` require a **running slideshow** (otherwise they return
`400 No slideshow running`).

Examples (from any machine on the network):

```bash
curl http://192.168.1.50:8000/ping
curl -X POST http://192.168.1.50:8000/next
curl -X POST http://192.168.1.50:8000/goto/5
curl -X POST -F "file=@deck.pptx" http://192.168.1.50:8000/open
```

Python:

```python
import requests
requests.post("http://192.168.1.50:8000/goto/3", timeout=5).json()
```

### Options

```bat
start.bat --port 8000 --upload-dir C:\Presentations
```

- `--port`: port to listen on (default `8000`).
- `--upload-dir`: where files sent via `/open` are saved (default `C:\Presentations`).

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Controller can't reach the laptop | Same Wi-Fi? Is the `start.bat` window open? Run `allow_firewall.bat` (accept the admin prompt). Some company/guest networks block device-to-device traffic: use a phone hotspot or your own router instead. Test from the other device with `curl http://<laptop-ip>:8000/ping`. |
| `No slideshow running` | Open the deck and press **F5**, or upload it via `/open`. |
| `setup.bat` fails while installing uv | Check the internet connection, or install uv manually: `winget install astral-sh.uv`, then run `setup.bat` again. |
| `setup.bat` says PowerPoint not detected | Install the desktop PowerPoint app (Microsoft 365 / Office). |
| Slides jump to the wrong number | Make sure it's the same `.pptx` and hidden slides haven't changed. |
| Something else | Read the error in the `start.bat` window. Restart `start.bat` and PowerPoint. |

---

## Security note

There is **no authentication**: anyone on the same network can change slides or upload
a `.pptx` to the laptop while `start.bat` is running. Only run it on trusted networks
and close it when you're done.

## Files

| File | Purpose |
|---|---|
| `setup.bat` | One-time setup: uv, Python, dependencies, firewall rule. |
| `start.bat` | Starts the server. |
| `allow_firewall.bat` | Opens TCP port 8000 in Windows Firewall (asks for admin). |
| `server.py` | The server itself (Flask + pywin32). |
| `pyproject.toml`, `uv.lock` | Dependencies, pinned for uv. |

Prefer to do it by hand? `uv sync` then `uv run server.py`, or with plain pip:
`pip install flask pywin32` then `python server.py`.
