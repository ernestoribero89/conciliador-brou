import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from flask import Flask, render_template, request, send_file, jsonify

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

BASE_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = BASE_DIR / "scripts"

SCRIPT_USD = SCRIPTS_DIR / "SCRIPT_USD_BROU.py"
SCRIPT_UYU = SCRIPTS_DIR / "SCRIPT_UYU_BROU.py"
SCRIPT_EUR = SCRIPTS_DIR / "SCRIPT_EUR_BROU.py"


@app.get("/")
def home():
    return render_template("index.html")


def save_upload(file_storage, path: Path):
    if not file_storage or file_storage.filename == "":
        raise ValueError(f"Falta archivo: {path.name}")
    file_storage.save(path)


def run_cmd(cmd, cwd):
    result = subprocess.run(
        cmd,
        cwd=cwd,
        text=True,
        capture_output=True,
        timeout=240
    )
    if result.returncode != 0:
        raise RuntimeError(
            "Falló el script:\n"
            + "COMANDO: " + " ".join(map(str, cmd)) + "\n\n"
            + "STDOUT:\n" + result.stdout + "\n\n"
            + "STDERR:\n" + result.stderr
        )
    return result.stdout


def conciliar_usd(work: Path, logs: list):
    sap = work / "SAP_BROU_USD.xlsx"
    banco = work / "BROU_USD.xlsx"

    stdout = run_cmd(["python", str(SCRIPT_USD), str(sap), str(banco)], cwd=work)
    logs.append(("USD", stdout))

    fixed_sap = work / "CONCILIACION_SAP.xlsx"
    fixed_bco = work / "CONCILIACION_BANCO.xlsx"

    out_sap = work / "CONCILIACION_SAP_USD.xlsx"
    out_bco = work / "CONCILIACION_BANCO_USD.xlsx"

    if fixed_sap.exists():
        fixed_sap.rename(out_sap)
    if fixed_bco.exists():
        fixed_bco.rename(out_bco)

    return [out_sap, out_bco]


def conciliar_uyu(work: Path, logs: list):
    sap = work / "SAP_BROU_UYU.xlsx"
    banco = work / "BROU_UYU.xlsx"

    out_sap = work / "CONCILIACION_SAP_UYU.xlsx"
    out_bco = work / "CONCILIACION_BANCO_UYU.xlsx"

    stdout = run_cmd(
        ["python", str(SCRIPT_UYU), str(sap), str(banco), str(out_sap), str(out_bco)],
        cwd=work
    )
    logs.append(("UYU", stdout))

    return [out_sap, out_bco]


def conciliar_eur(work: Path, logs: list):
    sap_uploaded = work / "SAP_BROU_EUR.xlsx"
    banco_uploaded = work / "BROU_EUR.xlsx"

    # El script EUR usa estos nombres internamente.
    shutil.copy(sap_uploaded, work / "sap_eur.xlsx")
    shutil.copy(banco_uploaded, work / "banco_eur.xlsx")

    stdout = run_cmd(["python", str(SCRIPT_EUR)], cwd=work)
    logs.append(("EUR", stdout))

    generated_sap = work / "SAP_EUR_conciliado.xlsx"
    generated_bco = work / "BANCO_EUR_conciliado.xlsx"

    out_sap = work / "CONCILIACION_SAP_EUR.xlsx"
    out_bco = work / "CONCILIACION_BANCO_EUR.xlsx"

    if generated_sap.exists():
        generated_sap.rename(out_sap)
    if generated_bco.exists():
        generated_bco.rename(out_bco)

    return [out_sap, out_bco]


@app.post("/conciliar")
def conciliar():
    try:
        with tempfile.TemporaryDirectory() as tmp:
            work = Path(tmp)
            logs = []
            outputs = []

            save_upload(request.files.get("sap_usd"), work / "SAP_BROU_USD.xlsx")
            save_upload(request.files.get("banco_usd"), work / "BROU_USD.xlsx")

            save_upload(request.files.get("sap_uyu"), work / "SAP_BROU_UYU.xlsx")
            save_upload(request.files.get("banco_uyu"), work / "BROU_UYU.xlsx")

            save_upload(request.files.get("sap_eur"), work / "SAP_BROU_EUR.xlsx")
            save_upload(request.files.get("banco_eur"), work / "BROU_EUR.xlsx")

            outputs += conciliar_usd(work, logs)
            outputs += conciliar_uyu(work, logs)
            outputs += conciliar_eur(work, logs)

            log_path = work / "RESUMEN_CONCILIACION.txt"
            with open(log_path, "w", encoding="utf-8") as f:
                for moneda, stdout in logs:
                    f.write("=" * 72 + "\n")
                    f.write(f"CONCILIACION {moneda}\n")
                    f.write("=" * 72 + "\n")
                    f.write(stdout)
                    f.write("\n\n")
            outputs.append(log_path)

            zip_path = work / "CONCILIACIONES_BROU.zip"
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
                for p in outputs:
                    if p.exists():
                        z.write(p, arcname=p.name)

            final_zip = Path(tempfile.gettempdir()) / "CONCILIACIONES_BROU.zip"
            shutil.copy(zip_path, final_zip)

            return send_file(
                final_zip,
                as_attachment=True,
                download_name="CONCILIACIONES_BROU.zip"
            )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True)
