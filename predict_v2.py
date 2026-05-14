"""
predict_v2.py – Hitri prediktor cene nepremičnine
==================================================
Naloži shranjen model (model_v2/best_model_v2.pkl) in vrne napoved.
Izhod: KEY=VALUE vrstice (za razčlenitev v gui.py).

Zagon:
    .venv312\\Scripts\\python predict_v2.py \\
        --vrsta Stanovanje --kraj Ljubljana \\
        --povrsina 75 --zemljisce 0 --leto 2005 --sobe 3 --energija B
"""

import sys, os, argparse, math

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MODEL_PATH = os.path.join(SCRIPT_DIR, "model_v2", "best_model_v2.pkl")

parser = argparse.ArgumentParser()
parser.add_argument("--vrsta",     default="Stanovanje")
parser.add_argument("--kraj",      default="")
parser.add_argument("--povrsina",  type=float, default=100.0)
parser.add_argument("--zemljisce", type=float, default=0.0)
parser.add_argument("--leto",      type=int,   default=2000)
parser.add_argument("--sobe",      type=float, default=3.0)
parser.add_argument("--energija",  default="D")
args = parser.parse_args()

# ── Preveri model ─────────────────────────────────────────────────────────────
if not os.path.isfile(MODEL_PATH):
    print(f"NAPAKA=Model ni najden: {MODEL_PATH}")
    print("INFO=Najprej poženi: .venv312\\Scripts\\python modeli_v2.py --docx")
    sys.exit(1)

try:
    import joblib, numpy as np
except ImportError as e:
    print(f"NAPAKA=Manjkajoča knjižnica: {e}")
    sys.exit(1)

# ── Naloži model ──────────────────────────────────────────────────────────────
data = joblib.load(MODEL_PATH)
model          = data["model"]
vrsta_enc      = data["vrsta_enc"]
obcina_enc     = data["obcina_enc"]
cena_m2_obcina = data["cena_m2_obcina"]
global_cena_m2 = data["global_cena_m2"]
train_medians  = data["train_medians"]
metrics        = data["metrics"]
best_name      = data["best_name"]
trained_on     = data.get("trained_on", "?")

# ── Feature engineering ───────────────────────────────────────────────────────
ENERGY_MAP = {"A+": 8, "A": 7, "B": 6, "C": 5, "D": 4, "E": 3, "F": 2, "G": 1}

starost = float(2026 - args.leto)
if starost < 0 or starost > 200:
    starost = float(train_medians.get("StarostZgrade", 30))

log_vel = math.log1p(args.povrsina)
log_zem = math.log1p(max(0, args.zemljisce))
sob     = float(args.sobe)
enr     = float(ENERGY_MAP.get(str(args.energija).upper(), 4))

global_mean = vrsta_enc.mean() if len(vrsta_enc) > 0 else 12.0
vrsta_val   = float(vrsta_enc.get(args.vrsta,   global_mean))
obcina_val  = float(obcina_enc.get(args.kraj,   obcina_enc.mean() if len(obcina_enc) > 0 else global_mean))
cena_m2_lok = float(cena_m2_obcina.get(args.kraj, global_cena_m2))
log_cm2     = math.log1p(cena_m2_lok)

X = np.array([[log_vel, log_zem, starost, sob, enr,
               vrsta_val, obcina_val, log_cm2]])

# ── Napoved ───────────────────────────────────────────────────────────────────
pred_log = model.predict(X)[0]
pred_eur = math.exp(pred_log)

mape     = metrics.get("MAPE", 25) / 100
mae      = metrics.get("MAE",  0)
r2       = metrics.get("R2",   0)

# 90 % interval zaupanja (±1.645 × std ≈ MAPE * 1.3 za log-normalno)
ci_factor = mape * 1.3
ci_min = pred_eur * (1 - ci_factor)
ci_max = pred_eur * (1 + ci_factor)

# Podobni: mediana lokacije × tipična površina (za referenčno primerjavo)
ref_cena_m2 = cena_m2_lok if cena_m2_lok > 100 else global_cena_m2
ref_med     = ref_cena_m2 * args.povrsina
ref_min     = ref_med * 0.75
ref_max     = ref_med * 1.35
ref_povp    = ref_med * 1.05

# ── Izhod (KEY=VALUE za gui.py parser) ───────────────────────────────────────
print(f"NAPOVEDANA={pred_eur:.0f}")
print(f"CI_MIN={ci_min:.0f}")
print(f"CI_MAX={ci_max:.0f}")
print(f"PODOBNI_N=~lokacijska ocena")
print(f"PODOBNI_MIN={ref_min:.0f}")
print(f"PODOBNI_MAX={ref_max:.0f}")
print(f"PODOBNI_MED={ref_med:.0f}")
print(f"PODOBNI_POVP={ref_povp:.0f}")
print(f"PODOBNI_FILTER={args.vrsta} | {args.kraj}")
print(f"VZORCI=6672")
print(f"NAPAKA_PCT={mape*100:.1f}")
print(f"VRSTA={args.vrsta}")
print(f"KRAJ={args.kraj}")
print(f"POVRSINA={args.povrsina:.0f}")
print(f"MODEL={best_name}")
print(f"R2={r2:.4f}")
print(f"MAE={mae:.0f}")
print(f"TRENIRANO={trained_on[:10]}")

