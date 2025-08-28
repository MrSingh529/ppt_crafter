from flask import Flask, request, send_file
import io, os, sys, tempfile, shutil, uuid, subprocess

app = Flask(__name__)

# Optional: GET /api/app -> quick health check
@app.get("/")
def health():
    return {"status": "ok"}

# POST /api/app  (Vercel mounts the function at /api/app)
DEFAULT_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "default_template.pptx")

@app.route("/", methods=["POST"])
def generate():
    has_excel = "excel" in request.files
    has_template = "template" in request.files and request.files["template"].filename

    if not has_excel:
        return ("Missing file: need 'excel'", 400)

    excel = request.files["excel"]
    ppt   = request.files["template"] if has_template else None

    if not excel.filename.lower().endswith((".xlsx", ".xls")):
        return ("Excel must be .xlsx or .xls", 400)
    if ppt and not ppt.filename.lower().endswith(".pptx"):
        return ("Template must be .pptx", 400)

    work = os.path.join(tempfile.gettempdir(), f"imarc_{uuid.uuid4().hex}")
    os.makedirs(work, exist_ok=True)

    try:
        excel_path = os.path.join(work, "datasheet_imarc.xlsx")
        excel.save(excel_path)

        # Use uploaded template if present and small enough, else fall back
        ppt_path = os.path.join(work, "template.pptx")
        if ppt:
            ppt.save(ppt_path)
        else:
            shutil.copyfile(DEFAULT_TEMPLATE_PATH, ppt_path)

        script_src = os.path.join(os.path.dirname(__file__), "generate_poc.py")
        script_dst = os.path.join(work, "generate_poc.py")
        shutil.copyfile(script_src, script_dst)

        proc = subprocess.run(
            [sys.executable, script_dst],
            cwd=work,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=120,
        )
        if proc.returncode != 0:
            return (f"Script failed\nSTDOUT:\n{proc.stdout}\n\nSTDERR:\n{proc.stderr}", 500)

        out_path = os.path.join(work, "updated_poc.pptx")
        if not os.path.exists(out_path):
            return ("Output PPTX not found (expected 'updated_poc.pptx')", 500)

        with open(out_path, "rb") as f:
            data = f.read()
        return send_file(
            io.BytesIO(data),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name="updated_poc.pptx",
        )
    except subprocess.TimeoutExpired:
        return ("Generation timed out. Try a smaller file or retry.", 504)
    except Exception as e:
        return (f"Server error: {e}", 500)
    finally:
        try: shutil.rmtree(work)
        except Exception: pass
