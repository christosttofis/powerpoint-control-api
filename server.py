"""
PowerPoint Control API — small Flask server for controlling PowerPoint on Windows.

Run on the Windows laptop that has PowerPoint installed. Any device on the same
network (phone, another PC, a script, an app) can then change slides over HTTP (default port 8000).

Easiest way: double-click setup.bat once, then start.bat every time.

Manual usage:
    python server.py [--port 8000] [--upload-dir C:\\Presentations]
"""

import os
import argparse
import functools
import socket
import threading

from flask import Flask, jsonify, request

app = Flask(__name__)

# Thread-local COM initialization
_com_lock = threading.Lock()


def com_initialized(func):
    """Decorator that ensures COM is initialized for the current thread."""

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        import pythoncom

        pythoncom.CoInitialize()
        try:
            return func(*args, **kwargs)
        finally:
            pythoncom.CoUninitialize()

    return wrapper


def get_powerpoint():
    """Get or create a PowerPoint Application COM object."""
    import win32com.client

    try:
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = True
    return ppt


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------


@app.route("/ping", methods=["GET"])
@com_initialized
def ping():
    """Health check — returns current slide info if PowerPoint is already running."""
    import win32com.client

    try:
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception:
        return jsonify({"ok": True, "slide": None, "message": "PowerPoint not running"})

    try:
        if ppt.Presentations.Count == 0:
            return jsonify({"ok": True, "slide": None, "message": "No presentation open"})

        pres = ppt.ActivePresentation
        slide = None
        try:
            ssw = ppt.SlideShowWindows
            if ssw.Count > 0:
                slide = ssw(1).View.Slide.SlideIndex
        except Exception:
            pass

        return jsonify({"ok": True, "slide": slide, "total_slides": pres.Slides.Count})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})


@app.route("/goto/<int:slide_number>", methods=["POST"])
@com_initialized
def goto_slide(slide_number):
    """Navigate to a specific slide in the running slideshow."""
    try:
        ppt = get_powerpoint()
        ssw = ppt.SlideShowWindows
        if ssw.Count == 0:
            return jsonify({"ok": False, "error": "No slideshow running"}), 400
        ssw(1).View.GotoSlide(slide_number)
        return jsonify({"ok": True, "slide": slide_number})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/next", methods=["POST"])
@com_initialized
def next_slide():
    """Advance to next slide."""
    try:
        ppt = get_powerpoint()
        ssw = ppt.SlideShowWindows
        if ssw.Count == 0:
            return jsonify({"ok": False, "error": "No slideshow running"}), 400
        view = ssw(1).View
        view.Next()
        slide = view.Slide.SlideIndex
        return jsonify({"ok": True, "slide": slide})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/prev", methods=["POST"])
@com_initialized
def prev_slide():
    """Go to previous slide."""
    try:
        ppt = get_powerpoint()
        ssw = ppt.SlideShowWindows
        if ssw.Count == 0:
            return jsonify({"ok": False, "error": "No slideshow running"}), 400
        view = ssw(1).View
        view.Previous()
        slide = view.Slide.SlideIndex
        return jsonify({"ok": True, "slide": slide})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/open", methods=["POST"])
@com_initialized
def open_presentation():
    """Upload and open a PPTX file in PowerPoint, then start the slideshow."""
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    f = request.files["file"]
    if not f.filename or not f.filename.endswith(".pptx"):
        return jsonify({"error": "Only .pptx files accepted"}), 400

    save_dir = app.config.get("UPLOAD_DIR", r"C:\Presentations")
    os.makedirs(save_dir, exist_ok=True)
    save_path = os.path.join(save_dir, os.path.basename(f.filename))
    f.save(save_path)

    try:
        ppt = get_powerpoint()
        # Close existing presentations
        while ppt.Presentations.Count > 0:
            ppt.Presentations(1).Close()
        # Open new one
        pres = ppt.Presentations.Open(os.path.abspath(save_path))
        pres.SlideShowSettings.Run()

        return jsonify(
            {
                "status": "opened",
                "filename": f.filename,
                "total_slides": pres.Slides.Count,
            }
        )
    except Exception as e:
        return jsonify({"error": f"Failed to open presentation: {e}"}), 500


def _local_ips():
    """Best-effort list of this laptop's IPv4 addresses (for clients to connect to)."""
    ips = set()
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ips.add(s.getsockname()[0])
        s.close()
    except Exception:
        pass
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ips.add(info[4][0])
    except Exception:
        pass
    return sorted(ip for ip in ips if not ip.startswith("127."))


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PowerPoint Control API")
    parser.add_argument("--port", type=int, default=8000, help="Port to listen on")
    parser.add_argument(
        "--upload-dir",
        default=r"C:\Presentations",
        help="Directory to save uploaded PPTX files",
    )
    args = parser.parse_args()

    app.config["UPLOAD_DIR"] = args.upload_dir
    print("=" * 60)
    print(f"[PowerPoint Control API] Starting on port {args.port}")
    print(f"[PowerPoint Control API] Upload directory: {args.upload_dir}")
    ips = _local_ips()
    if ips:
        print("[PowerPoint Control API] Clients should connect to one of:")
        for ip in ips:
            print(f"    {ip}:{args.port}")
    print("[PowerPoint Control API] Keep this window open during the presentation. Ctrl+C to stop.")
    print("=" * 60)
    app.run(host="0.0.0.0", port=args.port, threaded=True)
