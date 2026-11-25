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


def export_workbook_to_pdf(source_path, pdf_path, desktop):
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
        raise RuntimeError(f"LibreOffice nie otworzył pliku: {Path(source_path).name}")
    try:
        document.storeToURL(
            Path(pdf_path).resolve().as_uri(),
            (prop("FilterName", "calc_pdf_Export"), prop("Overwrite", True)),
        )
    finally:
        try:
            document.close(True)
        except Exception:
            pass
    if not Path(pdf_path).exists() or Path(pdf_path).stat().st_size < 100:
        raise RuntimeError(f"Nie udało się utworzyć PDF: {Path(pdf_path).name}")


def convert_zip_to_pdfs(zip_path, output_dir):
    max_files = 250
    max_uncompressed = 500 * 1024 * 1024
    with zipfile.ZipFile(zip_path) as archive:
        entries = [
            info for info in archive.infolist()
            if not info.is_dir() and info.filename.lower().endswith(".xlsx")
        ]
        if not entries:
            raise ValueError("Paczka ZIP nie zawiera plików XLSX.")
        if len(entries) > max_files:
            raise ValueError(f"Paczka zawiera zbyt wiele plików XLSX (limit: {max_files}).")
        total_size = sum(info.file_size for info in entries)
        if total_size > max_uncompressed:
            raise ValueError("Rozpakowana zawartość paczki przekracza limit 500 MB.")

        used_names = set()
        jobs = []
        input_dir = Path(output_dir) / "input"
        input_dir.mkdir(parents=True, exist_ok=True)
        for index, info in enumerate(entries, start=1):
            # Keep archive paths as PDF names, but never use them as filesystem paths.
            original_name = info.filename.replace("\\", "/")
            parts = [part for part in original_name.split("/") if part not in ("", ".")]
            if not parts or any(part == ".." for part in parts) or original_name.startswith("/"):
                raise ValueError("Paczka ZIP zawiera nieprawidłową ścieżkę pliku.")
            relative_pdf = str(Path(*parts).with_suffix(".pdf")).replace("\\", "/")
            if relative_pdf.casefold() in used_names:
                raise ValueError(f"W paczce są powtarzające się nazwy plików: {relative_pdf}")
            used_names.add(relative_pdf.casefold())

            input_path = input_dir / f"{index}_{safe_name(Path(parts[-1]).name)}"
            with archive.open(info) as source, open(input_path, "wb") as target:
                shutil.copyfileobj(source, target)
            pdf_path = Path(output_dir) / f"{index}.pdf"
            jobs.append((input_path, pdf_path, relative_pdf))

    with tempfile.TemporaryDirectory(prefix="lo-profile-") as profile_dir:
        process, desktop = connect_to_libreoffice(profile_dir)
        try:
            for source_path, pdf_path, _archive_name in jobs:
                export_workbook_to_pdf(source_path, pdf_path, desktop)
        finally:
            try:
                process.terminate()
                process.wait(timeout=8)
            except Exception:
                process.kill()

    return jobs

@app.get("/health")
def health():
    return jsonify({"status": "ok", "engine": "LibreOffice"})


@app.post("/api/pdf")
def create_pdf_archive():
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify({"error": "Nie przesłano paczki ZIP."}), 400
    if not upload.filename.lower().endswith(".zip"):
        return jsonify({"error": "Prześlij paczkę ZIP zawierającą pliki XLSX."}), 400

    with tempfile.TemporaryDirectory(prefix="xlsx-pdf-") as work_dir:
        source_path = Path(work_dir) / "source.zip"
        output_dir = Path(work_dir) / "output"
        output_dir.mkdir()
        upload.save(source_path)

        try:
            jobs = convert_zip_to_pdfs(source_path, output_dir)
            archive_path = Path(work_dir) / "pdfy.zip"
            with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as result:
                for _source, pdf_path, archive_name in jobs:
                    result.write(pdf_path, arcname=archive_name)

            return send_file(
                archive_path,
                mimetype="application/zip",
                as_attachment=True,
                download_name=f"{safe_name(Path(upload.filename).stem)}_PDF.zip",
            )
        except zipfile.BadZipFile:
            return jsonify({"error": "Przesłany plik nie jest prawidłową paczką ZIP."}), 400
        except Exception as error:
            app.logger.exception("PDF batch conversion failed")
            return jsonify({"error": f"Nie udało się wygenerować PDF: {error}"}), 500

@app.errorhandler(413)
def too_large(_error):
    return jsonify({"error": "Paczka ZIP jest za duża. Maksymalny rozmiar to 100 MB."}), 413
