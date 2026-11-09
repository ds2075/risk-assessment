"""
Веб-сервис оценки риска клиентов (НС ПОД/ФТ).

Пользователь загружает Excel-таблицу с вводными данными, сервер прогоняет её
через алгоритм risk_scoring.assess_file и отдаёт готовую таблицу с баллами,
уровнями риска и обоснованием.

Локальный запуск:  python app.py        → http://127.0.0.1:5000
На Render:         gunicorn app:app     (см. render.yaml)
"""

from __future__ import annotations

import os
import re
import tempfile
import time
import uuid
from collections import Counter
from pathlib import Path

from flask import Flask, abort, redirect, render_template, request, send_file, url_for
from werkzeug.exceptions import RequestEntityTooLarge

from risk_scoring import assess_file

BASE = Path(__file__).resolve().parent
WORK_DIR = Path(tempfile.gettempdir()) / "risk_results"
WORK_DIR.mkdir(parents=True, exist_ok=True)
KEEP_SECONDS = 60 * 60          # результаты хранятся 1 час, потом удаляются
MAX_UPLOAD_MB = 5

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

ID_RE = re.compile(r"^[0-9a-f]{32}$")


def cleanup_old_files() -> None:
    now = time.time()
    for f in WORK_DIR.glob("*.xlsx"):
        try:
            if now - f.stat().st_mtime > KEEP_SECONDS:
                f.unlink()
        except OSError:
            pass


def safe_stem(filename: str) -> str:
    stem = Path(filename).stem
    stem = re.sub(r'[\\/:*?"<>|]+', "_", stem).strip() or "таблица"
    return stem[:80]


@app.get("/")
def index():
    return render_template("index.html", max_mb=MAX_UPLOAD_MB)


@app.post("/process")
def process():
    cleanup_old_files()
    file = request.files.get("file")
    if not file or not file.filename:
        return render_template("index.html", error="Выберите файл с таблицей.", max_mb=MAX_UPLOAD_MB), 400
    if not file.filename.lower().endswith(".xlsx"):
        return render_template("index.html", max_mb=MAX_UPLOAD_MB,
                               error="Нужен файл Excel в формате .xlsx "
                                     "(в Excel: «Файл → Сохранить как → Книга Excel»)."), 400

    job = uuid.uuid4().hex
    src = WORK_DIR / f"{job}_in.xlsx"
    out = WORK_DIR / f"{job}.xlsx"
    file.save(src)
    try:
        results = assess_file(src, out)
    except ValueError as e:                       # понятные ошибки структуры таблицы
        return render_template("index.html", error=str(e), max_mb=MAX_UPLOAD_MB), 400
    except Exception:                             # битый файл, не Excel и т.п.
        app.logger.exception("Ошибка обработки файла")
        return render_template("index.html", max_mb=MAX_UPLOAD_MB,
                               error="Не удалось прочитать файл. Проверьте, что это таблица "
                                     "по шаблону, и попробуйте ещё раз."), 400
    finally:
        src.unlink(missing_ok=True)

    if not results:
        out.unlink(missing_ok=True)
        return render_template("index.html", max_mb=MAX_UPLOAD_MB,
                               error="В таблице не найдено ни одного клиента. "
                                     "Данные вводятся под строками заголовка."), 400

    counts = Counter(r.level for r in results)
    name = safe_stem(file.filename) + "_результат.xlsx"
    return render_template("result.html", results=results, counts=counts,
                           job=job, download_name=name, keep_min=KEEP_SECONDS // 60)


@app.get("/download/<job>")
def download(job: str):
    if not ID_RE.match(job):
        abort(404)
    path = WORK_DIR / f"{job}.xlsx"
    if not path.exists():
        return render_template("index.html", max_mb=MAX_UPLOAD_MB,
                               error="Срок хранения результата истёк — загрузите таблицу ещё раз."), 404
    name = request.args.get("name") or "результат.xlsx"
    return send_file(path, as_attachment=True, download_name=safe_stem(name) + ".xlsx")


@app.get("/template")
def template():
    return send_file(BASE / "static" / "template.xlsx", as_attachment=True,
                     download_name="Шаблон_таблицы_риска.xlsx")


@app.get("/example")
def example():
    return send_file(BASE / "static" / "example.xlsx", as_attachment=True,
                     download_name="Пример_клиентов.xlsx")


@app.get("/healthz")
def healthz():
    return "ok"


@app.errorhandler(RequestEntityTooLarge)
def too_large(_):
    return render_template("index.html", max_mb=MAX_UPLOAD_MB,
                           error=f"Файл больше {MAX_UPLOAD_MB} МБ."), 413


@app.errorhandler(404)
def not_found(_):
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
