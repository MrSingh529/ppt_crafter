from flask import Flask, request, send_file
import io, os, sys, tempfile, shutil, uuid, subprocess

app = Flask(__name__)

# Path to built-in default template (bundled in the function)
DEFAULT_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "default_template.pptx")

# Optional: GET /api/app -> quick health check
@app.get("/")
def health():
    return {"status": "ok"}

# POST /api/app  (Vercel mounts the function at /api/app)
@app.route("/", methods=["POST"])
def generate():
    # Excel is required; template is optional (falls back to default)
    if "excel" not in request.files:
        return ("Missing file: need 'excel'", 400)

    excel = request.files["excel"]
    ppt   = request.files.get("template")  # may be None or empty

    # Basic validation
    if not excel.filename.lower().endswith((".xlsx", ".xls")):
        return ("Excel must be .xlsx or .xls", 400)
    if ppt and ppt.filename and (not ppt.filename.lower().endswith(".pptx")):
        return ("Template must be .pptx", 400)

    # Ensure default exists when needed
    if (not ppt or not ppt.filename) and not os.path.exists(DEFAULT_TEMPLATE_PATH):
        return ("Server template missing. Please add api/default_template.pptx to the repo.", 500)

    work = os.path.join(tempfile.gettempdir(), f"imarc_{uuid.uuid4().hex}")
    os.makedirs(work, exist_ok=True)

    try:
        # Save uploads using the exact filenames your script expects
        excel_path = os.path.join(work, "datasheet_imarc.xlsx")
        ppt_path   = os.path.join(work, "template.pptx")

        excel.save(excel_path)
        if ppt and ppt.filename:
            ppt.save(ppt_path)  # use uploaded template
        else:
            shutil.copyfile(DEFAULT_TEMPLATE_PATH, ppt_path)  # use default template

        # Copy the *unchanged* business-logic script into the work dir
        script_src = os.path.join(os.path.dirname(__file__), "generate_poc.py")
        script_dst = os.path.join(work, "generate_poc.py")
        shutil.copyfile(script_src, script_dst)

        # Execute your script as-is
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
        try:
            shutil.rmtree(work)
        except Exception:
            pass
