# Conciliador BROU

App para subir 3 pares de archivos:

- SAP USD + Banco USD
- SAP UYU + Banco UYU
- SAP EUR + Banco EUR

Devuelve un ZIP con los Excel conciliados y un resumen.

## Estructura

```text
app.py
requirements.txt
README.md
templates/
  index.html
scripts/
  SCRIPT_USD_BROU.py
  SCRIPT_UYU_BROU.py
  SCRIPT_EUR_BROU.py
```

## Render

Build Command:

```text
pip install -r requirements.txt
```

Start Command:

```text
gunicorn app:app
```
