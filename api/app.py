import io, os, sys, tempfile, shutil, uuid, subprocess
from flask import Flask, request, send_file, abort

app = Flask(__name__)

# Route: POST /api/app.py/generate
@app.route("/generate", methods=["POST"])
def generate():
    if "excel" not in request.files or "template" not in request.files:
        return ("Missing files: need 'excel' and 'template'", 400)

    excel = request.files["excel"]
    ppt   = request.files["template"]

    # Basic validation
    if not excel.filename.lower().endswith((".xlsx", ".xls")):
        return ("Excel must be .xlsx or .xls", 400)
    if not ppt.filename.lower().endswith(".pptx"):
        return ("Template must be .pptx", 400)

    # Work dir in /tmp (required on Vercel)
    work = os.path.join(tempfile.gettempdir(), f"imarc_{uuid.uuid4().hex}")
    os.makedirs(work, exist_ok=True)

    try:
        # Save uploads with the exact filenames the script expects
        excel_path = os.path.join(work, "datasheet_imarc.xlsx")
        ppt_path   = os.path.join(work, "template.pptx")
        excel.save(excel_path)
        ppt.save(ppt_path)

        # Copy the *unchanged* script into the work dir
        script_src = os.path.join(os.path.dirname(__file__), "generate_poc.py")
        script_dst = os.path.join(work, "generate_poc.py")
        shutil.copyfile(script_src, script_dst)

        # Execute the script as-is
        proc = subprocess.run(
            [sys.executable, script_dst],
            cwd=work,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=120,
        )

        if proc.returncode != 0:
            # Propagate script logs for debugging
            return (f"Script failed\nSTDOUT:\n{proc.stdout}\n\nSTDERR:\n{proc.stderr}", 500)

        out_path = os.path.join(work, "updated_poc.pptx")
        if not os.path.exists(out_path):
            return ("Output PPTX not found (expected 'updated_poc.pptx')", 500)

        # Stream back as download
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
        # Best-effort cleanup
        try:
            shutil.rmtree(work)
        except Exception:
            pass