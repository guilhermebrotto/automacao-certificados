import os
import uuid
import zipfile
import threading
from pathlib import Path

from flask import Flask, request, jsonify, render_template, send_file, abort

from generate_certificates import (
    CertificateConfig,
    create_directories,
    read_csv_participants,
    generate_all_certificates,
    send_all_emails,
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

jobs = {}
email_jobs = {}


# ── helpers ───────────────────────────────────────────────────────────────────

def _make_job():
    return {
        "status": "running",
        "progress": 0.0,
        "current": 0,
        "total": 0,
        "messages": [],
        "report": None,
        "error": None,
        "participants": [],
        "config": None,
        "output_dir": None,
        "zip_path": None,
    }


def _update_progress(session_id, current, total, name):
    job = jobs.get(session_id)
    if job is None:
        return
    job["progress"] = round((current / total * 100) if total else 0, 1)
    job["current"] = current
    job["total"] = total
    job["messages"].append({"level": "info", "text": f"Gerando: {name} ({current}/{total})"})


# ── routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/gerar", methods=["POST"])
def gerar():
    csv_file = request.files.get("csv")
    pptx_file = request.files.get("pptx")
    event_name = request.form.get("event_name", "").strip()
    instructor_name = request.form.get("instructor_name", "").strip()

    errors = {}
    if not csv_file or not csv_file.filename:
        errors["csv"] = "Selecione o arquivo CSV de participantes."
    if not pptx_file or not pptx_file.filename:
        errors["pptx"] = "Selecione o modelo PPTX."
    if not event_name:
        errors["event_name"] = "Preencha o nome do evento."
    if not instructor_name:
        errors["instructor_name"] = "Preencha o nome do ministrante."
    if errors:
        return jsonify({"errors": errors}), 400

    session_id = str(uuid.uuid4())
    upload_dir = UPLOAD_DIR / session_id
    output_dir = OUTPUT_DIR / session_id
    upload_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    csv_path = upload_dir / "participantes.csv"
    pptx_path = upload_dir / "template.pptx"
    csv_file.save(str(csv_path))
    pptx_file.save(str(pptx_path))

    job = _make_job()
    job["output_dir"] = output_dir
    jobs[session_id] = job

    config = CertificateConfig(
        csv_file=csv_path,
        template_pptx=pptx_path,
        output_dir=output_dir,
        temp_dir=upload_dir / ".temp_pptx",
        event_name=event_name,
        instructor_name=instructor_name,
        progress_callback=lambda current, total, name: _update_progress(
            session_id, current, total, name
        ),
    )
    job["config"] = config

    threading.Thread(
        target=_run_generation, args=(session_id, config), daemon=True
    ).start()

    return jsonify({"session_id": session_id})


def _run_generation(session_id, config):
    job = jobs[session_id]
    try:
        create_directories(config)
        participants = read_csv_participants(config)

        if not participants:
            raise RuntimeError(
                "Nenhum participante encontrado no CSV. "
                "Verifique se o arquivo tem o formato correto: Nome;Email;Função"
            )

        job["total"] = len(participants)
        job["messages"].append(
            {"level": "info", "text": f"{len(participants)} participante(s) encontrado(s)."}
        )

        report = generate_all_certificates(participants, config)

        job["participants"] = participants
        job["report"] = report
        job["status"] = "done"
        job["progress"] = 100.0

        if report["failed"] > 0:
            job["messages"].append({
                "level": "warning",
                "text": (
                    f"Concluído: {report['success']}/{report['total']} gerados. "
                    f"{report['failed']} falha(s): {', '.join(report['failed_names'])}"
                ),
            })
        else:
            job["messages"].append({
                "level": "info",
                "text": f"Concluído com sucesso: {report['total']} certificado(s) gerado(s)!",
            })

    except Exception as exc:
        job["status"] = "error"
        job["error"] = str(exc)
        job["messages"].append({"level": "error", "text": str(exc)})


@app.route("/status/<session_id>")
def status(session_id):
    job = jobs.get(session_id)
    if job is None:
        abort(404)

    eligible = 0
    if job["status"] == "done":
        eligible = sum(
            1 for p in job["participants"]
            if p.get("email", "").endswith("@mindworks.com.br")
        )

    return jsonify({
        "status": job["status"],
        "progress": job["progress"],
        "current": job["current"],
        "total": job["total"],
        "messages": job["messages"],
        "report": job["report"],
        "error": job["error"],
        "eligible_emails": eligible,
    })


@app.route("/download/<session_id>")
def download(session_id):
    job = jobs.get(session_id)
    if job is None or job["status"] != "done":
        abort(404)

    zip_path = job.get("zip_path")
    if zip_path is None or not Path(zip_path).exists():
        output_dir = job["output_dir"]
        zip_path = OUTPUT_DIR / f"{session_id}.zip"
        files = list(output_dir.glob("*.pdf")) + list(output_dir.glob("*.pptx"))
        with zipfile.ZipFile(str(zip_path), "w", zipfile.ZIP_DEFLATED) as zf:
            for f in files:
                zf.write(str(f), f.name)
        job["zip_path"] = zip_path

    return send_file(str(zip_path), as_attachment=True, download_name="certificados.zip")


@app.route("/enviar/<session_id>", methods=["POST"])
def enviar(session_id):
    job = jobs.get(session_id)
    if job is None or job["status"] != "done":
        return jsonify({"error": "Job não encontrado ou geração não concluída."}), 400

    email_session_id = str(uuid.uuid4())
    email_job = {"status": "running", "messages": [], "report": None, "error": None}
    email_jobs[email_session_id] = email_job

    def _progress(current, total, name):
        email_job["messages"].append(
            {"level": "info", "text": f"Enviando para: {name} ({current}/{total})"}
        )

    threading.Thread(
        target=_run_email_sending,
        args=(email_session_id, job["participants"], job["config"], _progress),
        daemon=True,
    ).start()

    return jsonify({"email_session_id": email_session_id})


def _run_email_sending(email_session_id, participants, config, progress_callback):
    job = email_jobs[email_session_id]
    try:
        report = send_all_emails(participants, config, progress_callback)
        job["report"] = report
        job["status"] = "done"
        job["messages"].append({
            "level": "info",
            "text": (
                f"Envio concluído: {report['sent']} enviado(s), "
                f"{report['failed']} falha(s), {report['skipped']} ignorado(s)."
            ),
        })
    except Exception as exc:
        job["status"] = "error"
        job["error"] = str(exc)
        job["messages"].append({"level": "error", "text": str(exc)})


@app.route("/email-status/<email_session_id>")
def email_status(email_session_id):
    job = email_jobs.get(email_session_id)
    if job is None:
        abort(404)
    return jsonify(job)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
