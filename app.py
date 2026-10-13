import ast, json, os, threading
from datetime import datetime
from zoneinfo import ZoneInfo
import pandas as pd
from flask import Flask, jsonify, request, send_from_directory
from dotenv import load_dotenv
import os
#-------------------
load_dotenv()
# ----------------------------------------------------------------------------- Configuración
MODELO          = "gemini-3.5-flash-lite"
PRECIO_INPUT    = 0.30 / 1_000_000    # USD por token de entrada (Gemini 3.5 Flash-Lite)
PRECIO_OUTPUT   = 2.50 / 1_000_000  # USD por token de salida (Gemini 3.5 Flash-Lite)
TOPE_DIARIO_USD = float(os.getenv("TOPE_DIARIO_USD", "5"))
MAX_TOKENS_PLAN = 450                 # respuesta del plan de entrada
MAX_TOKENS_CHAT = 350                 # respuestas de seguimiento
MAX_TURNOS_CHAT = 6                   # historial que se reenvía (controla costo)
ZONA            = ZoneInfo("America/Bogota")
BASE            = os.path.dirname(os.path.abspath(__file__))
DATA            = os.path.join(BASE, "data")
USO_FILE        = os.path.join(DATA, "uso_diario.json")
CACHE_FILE      = os.path.join(DATA, "cache_planes.json")
API_KEY         = os.getenv("API_KEY")

app  = Flask(__name__, static_folder="static")
lock = threading.Lock()


# ----------------------------------------------------------------------------- Datos
def _lista(v):
    if isinstance(v, list): return v
    try: return ast.literal_eval(v) if isinstance(v, str) and v.startswith("[") else []
    except Exception: return []

emp = pd.read_csv(os.path.join(DATA, "scores_empresa_producto.csv"))
for c in ["google_comprados", "pp_ganadas", "pp_perdidas", "pp_abiertas"]:
    emp[c] = emp[c].apply(_lista)
emp = emp.where(pd.notna(emp), None)

cat_pp = pd.read_excel(os.path.join(DATA, "Cerebro_Producto_Propio_Servi.xlsx"))
cat_pp["nombre"] = cat_pp["Product Name"].str.split("(", regex=False).str[0].str.strip()
cat_pp = cat_pp.set_index("nombre")
cat_g = pd.read_excel(os.path.join(DATA, "Investigacion_Productos_Google.xlsx"))

def ficha_producto(nombre):
    if not nombre or nombre not in cat_pp.index: return "Sin ficha disponible."
    p = cat_pp.loc[nombre]
    return (f"Producto: {p['Product Name']}\nCategoría: {p['Category']}\n"
            f"Resumen: {str(p['Executive Summary (max 300 words)'])[:900]}\n"
            f"Roles compradores: {p['Buyer Roles (comma-separated, roles only)']}\n"
            f"Industrias objetivo: {p['Top Industries (comma-separated)']}\n"
            f"Tamaño mínimo: {p['Min Company Size']}\nComplejidad: {p['Complexity (Low/Medium/High)']}\n"
            f"Beneficio/ROI: {p['ROI/Main Benefit']}")

def contexto_cuenta(r):
    return (f"Empresa: {r['empresa']}\nMacro-industria: {r['MACRO']}\nPaís: {r['pais'] or 'N/D'}\n"
            f"Es cliente actual: {'Sí' if r['es_cliente'] else 'No'}\n"
            f"Productos Google que ya compra: {', '.join(r['google_comprados']) or 'ninguno'}\n"
            f"Producto propio que ya tiene: {', '.join(r['pp_ganadas']) or 'ninguno'}\n"
            f"Producto propio perdido antes: {', '.join(r['pp_perdidas']) or 'ninguno'}\n"
            f"Deal abierto: {'Sí' if r['deal_abierto'] else 'No'}\n"
            f"Score cuenta: {r['score_cuenta']}/100 · Prioridad: {r['prioridad']}\n"
            f"Producto sugerido: {r['producto_1']} (fit {r['fit_1']}) · Estado: {r['estado_1']}\n"
            f"Por qué: {r['razon_1']}\nSegunda opción: {r['producto_2'] or 'N/D'}")

SYSTEM = """Eres un asesor comercial senior de Servinformación (Información Localizada S.A.S.), empresa latinoamericana de soluciones de ubicación inteligente, datos e IA, y partner de Google Cloud.
Ayudas a un ejecutivo comercial a vender PRODUCTO PROPIO de Servinformación a una cuenta concreta.
Reglas:
- Usa solo los datos entregados. No inventes cifras, nombres de contactos ni casos de éxito.
- Español, tono profesional y directo. Sin emojis. Sin saludos ni cierres.
- Conecta lo que la cuenta ya compra en Google con el producto propio (expansión natural de la relación).
- Si el producto se perdió antes, reconócelo y sugiere un ángulo distinto."""

FORMATO_PLAN = """Entrega un plan de entrada breve (máximo 170 palabras) con exactamente estas secciones en markdown:
**Lectura de la cuenta** (1-2 frases)
**Ángulo de entrada** (1-2 frases: el dolor probable de su industria y cómo lo resuelve el producto)
**A quién buscar** (2-3 roles)
**Pregunta de apertura** (una pregunta textual para la primera reunión)
**Objeción probable y respuesta** (1 frase cada una)
**Siguiente paso** (una acción concreta)"""

# ----------------------------------------------------------------------------- Control de gasto
def _hoy(): return datetime.now(ZONA).strftime("%Y-%m-%d")

def _leer(path, default):
    try:
        with open(path, encoding="utf-8") as f: return json.load(f)
    except Exception: return default

def _guardar(path, obj):
    with open(path, "w", encoding="utf-8") as f: json.dump(obj, f, ensure_ascii=False)

def gasto_hoy():
    return _leer(USO_FILE, {}).get(_hoy(), {}).get("usd", 0.0)
def llamar_claude(mensajes, max_tokens):
    """Reserva el peor caso antes de llamar; si excede el tope, no llama."""
    if not API_KEY:
        raise RuntimeError("Falta configurar GOOGLE_API_KEY en el servidor.")
    tokens_in_est = sum(len(m["content"]) for m in mensajes) // 3 + len(SYSTEM) // 3
    peor_caso = tokens_in_est * PRECIO_INPUT + max_tokens * PRECIO_OUTPUT
    with lock:
        uso = _leer(USO_FILE, {})
        dia = uso.setdefault(_hoy(), {"usd": 0.0, "llamadas": 0})
        if dia["usd"] + peor_caso > TOPE_DIARIO_USD:
            raise PermissionError("Se alcanzó el límite de consumo de hoy. El asistente vuelve a estar disponible mañana.")
        dia["usd"] += peor_caso; _guardar(USO_FILE, uso)
    
    try:
        import requests
        
        # URL para Google Cloud Generative AI (usa la clave como API key en header)
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{MODELO}:generateContent"
        
        headers = {
            "Content-Type": "application/json",
            "x-goog-api-key": API_KEY
        }
        
        # Convertir mensajes a formato Gemini
        contents = []
        for m in mensajes:
            contents.append({
                "role": "user" if m["role"] == "user" else "model",
                "parts": [{"text": m["content"]}]
            })
        
        payload = {
            "contents": contents,
            "systemInstruction": {"parts": [{"text": SYSTEM}]},
            "generationConfig": {"maxOutputTokens": max_tokens}
        }
        
        r = requests.post(url, json=payload, headers=headers, timeout=60)
        r.raise_for_status()
        data = r.json()
        
        # Extraer texto
        texto = ""
        if "candidates" in data and len(data["candidates"]) > 0:
            candidate = data["candidates"][0]
            if "content" in candidate and "parts" in candidate["content"]:
                for part in candidate["content"]["parts"]:
                    if "text" in part:
                        texto += part["text"]
        
        # Extraer tokens
        usage = data.get("usageMetadata", {})
        input_tokens = usage.get("promptTokenCount", tokens_in_est)
        output_tokens = usage.get("candidatesTokenCount", len(texto) // 3)
        real = input_tokens * PRECIO_INPUT + output_tokens * PRECIO_OUTPUT
        
    except Exception as e:
        print(f"Error: {e}")
        texto, real = "", 0.0
    
    with lock:
        uso = _leer(USO_FILE, {}); dia = uso[_hoy()]
        dia["usd"] = round(dia["usd"] - peor_caso + real, 6); dia["llamadas"] += 1 if texto else 0
        _guardar(USO_FILE, uso)
    
    if not texto: raise RuntimeError("No fue posible contactar el modelo. Intenta de nuevo.")
    return texto
# ----------------------------------------------------------------------------- API
@app.get("/")
def index(): return send_from_directory("static", "index.html")

@app.get("/api/ejecutivos")
def ejecutivos():
    g = emp[emp.owner.notna() & (emp.owner != "(Sin valor)")].groupby("owner").agg(cuentas=("emp", "count"), alta=("prioridad", lambda s: int((s == "Alta").sum())))
    return jsonify([{"nombre": k, **v} for k, v in g.sort_index().to_dict("index").items()])

@app.get("/api/cuentas")
def cuentas():
    o = request.args.get("owner", "")
    d = emp[emp.owner == o].sort_values("score_final", ascending=False)
    cols = ["emp", "empresa", "MACRO", "pais", "prioridad", "score_final", "score_cuenta", "producto_1", "fit_1",
            "estado_1", "razon_1", "producto_2", "google_comprados", "pp_ganadas", "es_cliente", "deal_abierto",
            "pts_cliente", "pts_penetracion", "pts_whitespace", "pts_presupuesto", "pts_deal_abierto"]
    return jsonify(json.loads(d[cols].to_json(orient="records", force_ascii=False)))

@app.get("/api/uso")
def uso(): return jsonify({"usd": round(gasto_hoy(), 3), "tope": TOPE_DIARIO_USD})

def _fila(emp_id):
    d = emp[emp.emp == int(emp_id)]
    if d.empty: raise KeyError
    return d.iloc[0]

@app.post("/api/plan")
def plan():
    emp_id = str(request.json.get("emp"))
    cache = _leer(CACHE_FILE, {})
    if emp_id in cache: return jsonify({"texto": cache[emp_id], "cache": True})
    try:
        r = _fila(emp_id)
        prompt = f"DATOS DE LA CUENTA\n{contexto_cuenta(r)}\n\nFICHA DEL PRODUCTO PROPIO SUGERIDO\n{ficha_producto(r['producto_1'])}\n\n{FORMATO_PLAN}"
        texto = llamar_claude([{"role": "user", "content": prompt}], MAX_TOKENS_PLAN)
    except KeyError: return jsonify({"error": "Cuenta no encontrada."}), 404
    except PermissionError as e: return jsonify({"error": str(e)}), 429
    except RuntimeError as e: return jsonify({"error": str(e)}), 503
    cache[emp_id] = texto; _guardar(CACHE_FILE, cache)
    return jsonify({"texto": texto, "cache": False})

@app.post("/api/chat")
def chat():
    body = request.json or {}
    try:
        r = _fila(body.get("emp"))
        hist = [m for m in body.get("historial", []) if m.get("role") in ("user", "assistant")][-MAX_TURNOS_CHAT:]
        pregunta = str(body.get("pregunta", ""))[:600]
        base = (f"DATOS DE LA CUENTA\n{contexto_cuenta(r)}\n\nFICHA DEL PRODUCTO\n{ficha_producto(r['producto_1'])}\n\n"
                "Responde preguntas del ejecutivo sobre esta cuenta en máximo 120 palabras.")
        msgs = [{"role": "user", "content": base}, {"role": "assistant", "content": "Entendido."}] + hist + \
               [{"role": "user", "content": pregunta}]
        texto = llamar_claude(msgs, MAX_TOKENS_CHAT)
    except KeyError: return jsonify({"error": "Cuenta no encontrada."}), 404
    except PermissionError as e: return jsonify({"error": str(e)}), 429
    except RuntimeError as e: return jsonify({"error": str(e)}), 503
    return jsonify({"texto": texto})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8000")))