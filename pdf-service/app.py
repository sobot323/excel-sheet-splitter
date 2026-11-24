import os
import re
import shutil
import socket
import subprocess
import tempfile
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS

# Debian's python3-uno module is installed outside the /usr/local Python path.
import sys
sys.path.append("/usr/lib/python3/dist-packages")
import uno
from com.sun.star.beans import PropertyValue

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_MB", "100")) * 1024 * 1024

allowed_origin = os.getenv("FRONTEND_ORIGIN", "https://sobot323.github.io")
CORS(app, resources={r"/api/*": {"origins": [allowed_origin]}})


def prop(name, value):
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def safe_name(value):
    value = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", str(value or "Arkusz"))
    value = re.sub(r"[. ]+$", "", value).strip()
    return value or "Arkusz"


def unique_name(name, used):
    candidate = name
    number = 2
    while candidate.casefold() in used:
        candidate = f"{name}_{number}"
        number += 1
    used.add(candidate.casefold())
    return candidate


def get_sheet_names(xlsx_path):
    # Read sheet names from the workbook package without modifying the source file.
    with zipfile.ZipFile(xlsx_path) as archive:
        root = ET.fromstring(archive.read("xl/workbook.xml"))
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    names = [node.attrib["name"] for node in root.findall(".//m:sheets/m:sheet", ns)]
    if not names:
        raise ValueError("Skoroszyt nie zawiera arkuszy.")
    return names


def free_port():
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def connect_to_libreoffice(profile_dir):
    port = free_port()
    accept = f"socket,host=127.0.0.1,port={port};urp;StarOffice.ServiceManager"
    command = [
        "soffice", "--headless", "--nologo", "--nodefault", "--nofirststartwizard",
        f"-env:UserInstallation={Path(profile_dir).as_uri()}",
        f"--accept={accept}",
    ]
    process = subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )
        connection = f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
        deadline = time.monotonic() + 35
        while time.monotonic() < deadline:
            if process.poll() is not None:
                raise RuntimeError("LibreOffice zakończył działanie podczas uruchamiania.")
            try:
                context = resolver.resolve(connection)
                desktop = context.ServiceManager.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", context
                )
                return process, desktop
            except Exception:
                time.sleep(0.25)
        raise RuntimeError("Nie udało się uruchomić silnika LibreOffice.")
    except Exception:
        process.terminate()
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            process.kill()
        raise


def export_sheets_to_pdf(source_path, output_dir, names, suffix):
    with tempfile.TemporaryDirectory(prefix="lo-profile-") as profile_dir:
        process, desktop = connect_to_libreoffice(profile_dir)
        document = None
        try:
            source_url = Path(source_path).resolve().as_uri()
            document = desktop.loadComponentFromURL(
                source_url,
                "_blank",
                0,
                (
                    prop("Hidden", True),
                    prop("ReadOnly", True),
                    prop("UpdateDocMode", 3),
                    prop("MacroExecutionMode", 0),
                ),
            )
            if document is None:
                raise RuntimeError("LibreOffice nie otworzył skoroszytu.")

            sheets = document.getSheets()
            actual_names = list(sheets.getElementNames())
            if actual_names != names:
                # Use the workbook's actual sheet order and names as read by LibreOffice.
                names = actual_names
            used = set()
            output_files = []

            for selected_name in names:
                # Keep exactly one visible sheet so the PDF export contains only that tab.
                for sheet_name in names:
                    sheets.getByName(sheet_name).IsVisible = (sheet_name == selected_name)
                document.getCurrentController().setActiveSheet(sheets.getByName(selected_name))

                filename = unique_name(safe_name(selected_name + suffix), used) + ".pdf"
                pdf_path = Path(output_dir) / filename
                # Export only the selected sheet range. Hiding other sheets alone
                # is not sufficient because Calc's PDF export can still include them.
                selected_sheet = sheets.getByName(selected_name)
                filter_data = (prop("Selection", selected_sheet),)
                document.storeToURL(
                    pdf_path.resolve().as_uri(),
                    (
                        prop("FilterName", "calc_pdf_Export"),
                        prop("FilterData", filter_data),
                        prop("Overwrite", True),
                    ),
                )
                if not pdf_path.exists() or pdf_path.stat().st_size < 500:
                    raise RuntimeError(f"Nie udało się utworzyć PDF dla arkusza: {selected_name}")
                output_files.append(pdf_path)

            return output_files
        finally:
            if document is not None:
                try:
                    document.close(True)
                except Exception:
                    pass
            try:
                process.terminate()
                process.wait(timeout=8)
            except Exception:
                process.kill()


@app.get("/health")
def health():
    return jsonify({"status": "ok", "engine": "LibreOffice"})


@app.post("/api/pdf")
def create_pdf_archive():
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify({"error": "Nie przesłano pliku XLSX."}), 400
    if not upload.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "Obsługiwany jest wyłącznie format XLSX."}), 400

    suffix = request.form.get("suffix", "")
    if len(suffix) > 100:
        return jsonify({"error": "Tekst dodawany do nazw plików jest za długi."}), 400

    with tempfile.TemporaryDirectory(prefix="xlsx-pdf-") as work_dir:
        source_path = Path(work_dir) / "source.xlsx"
        output_dir = Path(work_dir) / "pdf"
        output_dir.mkdir()
        upload.save(source_path)

        try:
            names = get_sheet_names(source_path)
            pdf_files = export_sheets_to_pdf(source_path, output_dir, names, suffix)
            archive_path = Path(work_dir) / "pdfy.zip"
            with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as archive:
                for pdf_path in pdf_files:
                    archive.write(pdf_path, arcname=pdf_path.name)

            return send_file(
                archive_path,
                mimetype="application/zip",
                as_attachment=True,
                download_name=f"{safe_name(Path(upload.filename).stem)}_PDF.zip",
            )
        except Exception as error:
            app.logger.exception("PDF export failed")
            return jsonify({"error": f"Nie udało się wygenerować PDF: {error}"}), 500


@app.errorhandler(413)
def too_large(_error):
    return jsonify({"error": "Plik jest za duży. Maksymalny rozmiar to 100 MB."}), 413
