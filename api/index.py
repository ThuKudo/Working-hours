import json
import sys
from pathlib import Path

from flask import Flask, Response, jsonify, request


ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from webapp import HTML_PAGE, STORE_DATA, ensure_contacts_workbook, save_contact  # noqa: E402


app = Flask(__name__)


@app.route("/", defaults={"path": ""}, methods=["GET", "POST"])
@app.route("/<path:path>", methods=["GET", "POST"])
def router(path: str):
    current_path = "/" + path

    if request.method == "GET" and current_path == "/":
        html = HTML_PAGE.replace("__STORE_DATA__", json.dumps(STORE_DATA, ensure_ascii=False))
        return Response(html, mimetype="text/html; charset=utf-8")

    if request.method == "GET" and current_path == "/api/storefront":
        return jsonify(STORE_DATA)

    if request.method == "POST" and current_path == "/api/contact":
        payload = request.get_json(silent=True) or {}
        name = str(payload.get("name", "")).strip()
        phone = str(payload.get("phone", "")).strip()
        if not name or not phone:
            return jsonify({"error": "Name and phone are required."}), 400

        ensure_contacts_workbook()
        save_contact(
            {
                "form_type": str(payload.get("form_type", "")).strip(),
                "name": name,
                "phone": phone,
                "message": str(payload.get("message", "")).strip(),
                "language": str(payload.get("language", "")).strip(),
                "source": str(payload.get("source", "")).strip(),
            }
        )
        return jsonify({"status": "ok"}), 201

    return jsonify({"error": "Not found"}), 404
