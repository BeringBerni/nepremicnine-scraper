"""
cenik.py – Napovednik cen nepremičnin  (s predpomnjenjem modela)
================================================================
Model se samodejno shrani po prvem treningu.  Pri naslednjih klicih
se le naloži iz diska → napoved je takojšnja (~0.1 s namesto ~10 s).

Predpomnilnik:  <ime_csv>_cenik_rf<n>_d<depth>.pkl   (poleg CSV datoteke)

Zagon:
    py cenik.py --vrsta Samostojna --kraj Kranj --povrsina 150
    py cenik.py --retrain           # Prisili ponovni trening
    py cenik.py --samo-trening      # Samo natrenira + shrani, brez napovedi
"""

import sys, os, csv, math, random, statistics, argparse, pickle
from collections import defaultdict

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.setrecursionlimit(50_000)

# ── CLI argumenti ─────────────────────────────────────────────────────────────
parser = argparse.ArgumentParser(description="Napovednik cen nepremičnin")
parser.add_argument("--csv",          default=None,  help="Pot do CSV datoteke")
parser.add_argument("--vrsta",        default=None,  help="Vrsta objekta (npr. Samostojna)")
parser.add_argument("--kraj",         default=None,  help="Občina / kraj (npr. Kranj)")
parser.add_argument("--povrsina",     default=None,  type=float, help="Površina v m²")
parser.add_argument("--zemljisce",    default=None,  type=float, help="Površina zemljišča v m²")
parser.add_argument("--leto",         default=None,  type=float, help="Leto gradnje")
parser.add_argument("--sobe",         default=None,  type=float, help="Število sob")
parser.add_argument("--energija",     default=None,  help="Energetski razred (A–G)")
parser.add_argument("--n-trees",      default=60,    type=int,   dest="n_trees")
parser.add_argument("--depth",        default=7,     type=int)
parser.add_argument("--seed",         default=42,    type=int)
parser.add_argument("--retrain",      action="store_true", help="Prisili ponovni trening modela")
parser.add_argument("--samo-trening", action="store_true", dest="samo_trening",
                    help="Samo natrenira in shrani model (brez napovedi)")
args = parser.parse_args()

random.seed(args.seed)

# ── Iskanje CSV ────────────────────────────────────────────────────────────────
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

CANDIDATES = [
    args.csv,
    os.path.join(_SCRIPT_DIR, "nepremicnine_export_prodaja.csv"),
    os.path.join(_SCRIPT_DIR, "nepremicnine_export_najem.csv"),
    os.path.join(_SCRIPT_DIR, "nepremicnine_export.csv"),
    os.path.join(_SCRIPT_DIR, "UUI-lv2", "bin", "Debug", "net10.0", "nepremicnine_export.csv"),
    "nepremicnine_export_prodaja.csv",
    "nepremicnine_export_najem.csv",
    "nepremicnine_export.csv",
]
csv_path = next((p for p in CANDIDATES if p and os.path.isfile(p)), None)
if csv_path is None:
    print("NAPAKA=CSV datoteka ni najdena")
    sys.exit(1)

csv_path = os.path.abspath(csv_path)

# ── Pot do predpomnilnika ──────────────────────────────────────────────────────
_csv_base  = os.path.splitext(csv_path)[0]
cache_path = f"{_csv_base}_cenik_rf{args.n_trees}_d{args.depth}.pkl"

# ── Energetski razred → število ────────────────────────────────────────────────
_ENERGY = {"A+": 8, "A": 7, "B": 6, "C": 5, "D": 4, "E": 3, "F": 2, "G": 1}


# ──────────────────────────────────────────────────────────────────────────────
#  MODEL: Decision Tree + Random Forest  (čista standardna knjižnica)
# ──────────────────────────────────────────────────────────────────────────────
class _Node:
    __slots__ = ("feat", "thresh", "left", "right", "val")
    def __init__(self, **kw):
        for k, v in kw.items():
            setattr(self, k, v)


class _DT:
    def __init__(self, max_depth=7, min_s=4, max_f=None):
        self.max_depth = max_depth
        self.min_s     = min_s
        self.max_f     = max_f

    def fit(self, X, y):
        self._root = self._build(X, y, 0)
        return self

    def _split(self, X, y):
        n, p        = len(X), len(X[0])
        bst, bf, bt = -1.0, None, None
        tv          = sum(v ** 2 for v in y) / n - (sum(y) / n) ** 2
        feats       = list(range(p))
        if self.max_f and self.max_f < p:
            feats = random.sample(feats, self.max_f)
        for fi in feats:
            order      = sorted(range(n), key=lambda i: X[i][fi])
            ps, ps2    = [0.0] * (n + 1), [0.0] * (n + 1)
            for k in range(n):
                i = order[k]; ps[k+1] = ps[k]+y[i]; ps2[k+1] = ps2[k]+y[i]**2
            for k in range(1, n):
                if X[order[k]][fi] == X[order[k-1]][fi]:
                    continue
                nl, nr  = k, n - k
                if nl < 2 or nr < 2:
                    continue
                sl, s2l = ps[k], ps2[k]
                sr, s2r = ps[n] - sl, ps2[n] - s2l
                g = tv - (nl * (s2l/nl - (sl/nl)**2) + nr * (s2r/nr - (sr/nr)**2)) / n
                if g > bst:
                    bst = g; bf = fi
                    bt  = (X[order[k-1]][fi] + X[order[k]][fi]) / 2
        return bf, bt

    def _build(self, X, y, d):
        val = sum(y) / len(y)
        if len(y) < self.min_s or (self.max_depth and d >= self.max_depth):
            return _Node(val=val, feat=None, thresh=None, left=None, right=None)
        f, t = self._split(X, y)
        if f is None:
            return _Node(val=val, feat=None, thresh=None, left=None, right=None)
        mask = [X[i][f] <= t for i in range(len(X))]
        Xl   = [X[i] for i in range(len(X)) if     mask[i]]
        yl   = [y[i] for i in range(len(y)) if     mask[i]]
        Xr   = [X[i] for i in range(len(X)) if not mask[i]]
        yr   = [y[i] for i in range(len(y)) if not mask[i]]
        if not Xl or not Xr:
            return _Node(val=val, feat=None, thresh=None, left=None, right=None)
        return _Node(val=val, feat=f, thresh=t,
                     left=self._build(Xl, yl, d+1),
                     right=self._build(Xr, yr, d+1))

    def _p1(self, nd, x):
        if nd.feat is None:
            return nd.val
        return self._p1(nd.left if x[nd.feat] <= nd.thresh else nd.right, x)

    def predict(self, X):
        return [self._p1(self._root, x) for x in X]


class RandomForest:
    def __init__(self, n=60, depth=7, min_s=4):
        self.n = n; self.depth = depth; self.min_s = min_s
        self._trees = []

    def fit(self, X, y):
        n, p        = len(X), len(X[0])
        mf          = max(1, int(p ** 0.5))
        self._trees = []
        print(f"   Treniram model ({self.n} dreves, globina {self.depth})…", flush=True)
        for i in range(self.n):
            idx = [random.randint(0, n-1) for _ in range(n)]
            t   = _DT(max_depth=self.depth, min_s=self.min_s, max_f=mf)
            t.fit([X[j] for j in idx], [y[j] for j in idx])
            self._trees.append(t)
            if (i + 1) % 20 == 0:
                print(f"   {i+1}/{self.n} dreves", flush=True)
        return self

    def predict_all(self, x):
        """Vrne napovedi vseh dreves za en vzorec (za CI izračun)."""
        return [t._p1(t._root, x) for t in self._trees]

    def predict(self, X):
        return [statistics.mean(t.predict([x])[0] for t in self._trees) for x in X]


# ──────────────────────────────────────────────────────────────────────────────
#  PREDPOMNILNIK  (dva ločena pickle objekta v eni datoteki: meta + model)
# ──────────────────────────────────────────────────────────────────────────────
def _cache_meta(path):
    """Hitro naloži samo metapodatke (1. objekt). Vrne dict ali None."""
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "rb") as f:
            return pickle.load(f)
    except Exception:
        return None


def _cache_valid(path, csv_path, n_trees, depth, seed):
    """Vrni True, če je predpomnilnik svež in parametri se ujemajo."""
    m = _cache_meta(path)
    if not m:
        return False
    return (
        m.get("csv_mtime") == os.path.getmtime(csv_path) and
        m.get("csv_size")  == os.path.getsize(csv_path)  and
        m.get("n_trees")   == n_trees and
        m.get("depth")     == depth   and
        m.get("seed")      == seed
    )


def _load_cache_model(path):
    """Naloži 2. pickle objekt (dejanski model)."""
    with open(path, "rb") as f:
        pickle.load(f)           # preskočimo meta
        return pickle.load(f)    # ← model


def _save_cache(path, meta, model_data):
    """Shrani: 1. meta (majhen, za hitro validacijo)  2. model (velik)."""
    with open(path, "wb") as f:
        pickle.dump(meta,       f, protocol=pickle.HIGHEST_PROTOCOL)
        pickle.dump(model_data, f, protocol=pickle.HIGHEST_PROTOCOL)


# ──────────────────────────────────────────────────────────────────────────────
#  NALAGANJE CSV + PREDPROCESIRANJE + TRENING
# ──────────────────────────────────────────────────────────────────────────────
def _f(s):
    try:
        v = float(str(s).replace(",", ".").strip())
        return v if math.isfinite(v) else None
    except Exception:
        return None


def _med(rows, key):
    vals = [r[key] for r in rows if r[key] is not None]
    return statistics.median(vals) if vals else 0.0


def load_and_train(csv_path, n_trees, depth, seed):
    """Naloži CSV, predprocesiraj in natrenira model. Vrni (meta, model_data)."""
    raw_rows = []
    with open(csv_path, newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh, delimiter=";"):
            cena = _f(r.get("Cena"))
            if cena is None or not (5_000 < cena < 5_000_000):
                continue
            m2   = _f(r.get("VelikostM2"))
            zem  = _f(r.get("ZemljisteM2"))
            leto = _f(r.get("LetoGradnje"))
            sob  = _f(r.get("StSob"))
            enr  = _ENERGY.get(str(r.get("EnergetskiRazred", "")).strip().upper(), None)
            if m2 is not None and m2 >= 5_000:
                m2 = None
            raw_rows.append({
                "Cena":          cena,
                "VelikostM2":    m2,
                "ZemljisteM2":   zem,
                "LetoGradnje":   leto if (leto and leto > 1_900) else None,
                "StSob":         sob,
                "EnergetRazred": float(enr) if enr else None,
                "VrstaObjekta":  r.get("VrstaObjekta", "").strip(),
                "Obcina":        (r.get("Obcina") or r.get("Lokacija", "")).strip(),
            })

    n_vzorcev = len(raw_rows)
    if n_vzorcev < 20:
        print("NAPAKA=Premalo podatkov (< 20 vrstic)")
        sys.exit(1)

    # Imputacija
    imp = {k: _med(raw_rows, k) for k in
           ["VelikostM2", "ZemljisteM2", "LetoGradnje", "StSob", "EnergetRazred"]}

    # VrstaObjekta → koda po frekvenci
    vc = defaultdict(int)
    for r in raw_rows:
        vc[r["VrstaObjekta"]] += 1
    vrsta_enc = {v: i for i, (v, _) in
                 enumerate(sorted(vc.items(), key=lambda x: -x[1]))}

    # Obcina → ordinalna koda po mediani cene
    ob_pr = defaultdict(list)
    for r in raw_rows:
        if r["Obcina"]:
            ob_pr[r["Obcina"]].append(r["Cena"])
    ob_sorted  = sorted(ob_pr, key=lambda k: statistics.median(ob_pr[k]))
    obcina_enc = {k: i for i, k in enumerate(ob_sorted)}

    def row_to_x(r):
        return [
            r["VelikostM2"]    if r["VelikostM2"]    is not None else imp["VelikostM2"],
            r["ZemljisteM2"]   if r["ZemljisteM2"]   is not None else imp["ZemljisteM2"],
            r["LetoGradnje"]   if r["LetoGradnje"]    is not None else imp["LetoGradnje"],
            r["StSob"]         if r["StSob"]          is not None else imp["StSob"],
            r["EnergetRazred"] if r["EnergetRazred"]  is not None else imp["EnergetRazred"],
            float(vrsta_enc.get(r["VrstaObjekta"], 0)),
            float(obcina_enc.get(r["Obcina"], len(obcina_enc) // 2)),
        ]

    all_X = [row_to_x(r) for r in raw_rows]
    all_y = [r["Cena"] for r in raw_rows]

    # Standardizacija
    p       = len(all_X[0])
    f_means = [statistics.mean(row[j] for row in all_X) for j in range(p)]
    f_stds  = [(statistics.stdev(row[j] for row in all_X) or 1.0) for j in range(p)]
    X_s     = [[(row[j] - f_means[j]) / f_stds[j] for j in range(p)] for row in all_X]

    # Trening
    random.seed(seed)
    rf = RandomForest(n=n_trees, depth=depth).fit(X_s, all_y)

    meta = {
        "csv_mtime": os.path.getmtime(csv_path),
        "csv_size":  os.path.getsize(csv_path),
        "n_trees":   n_trees,
        "depth":     depth,
        "seed":      seed,
    }
    model_data = {
        "rf":         rf,
        "f_means":    f_means,
        "f_stds":     f_stds,
        "imp":        imp,
        "vrsta_enc":  vrsta_enc,
        "obcina_enc": obcina_enc,
        "ob_sorted":  ob_sorted,
        "vc":         dict(vc),
        "raw_rows":   raw_rows,
        "n_vzorcev":  n_vzorcev,
    }
    return meta, model_data


# ──────────────────────────────────────────────────────────────────────────────
#  NALOŽI ALI NATRENIRA
# ──────────────────────────────────────────────────────────────────────────────
use_cache = (
    not args.retrain and
    not args.samo_trening and
    _cache_valid(cache_path, csv_path, args.n_trees, args.depth, args.seed)
)

if use_cache:
    print(f"✓  Nalagam model iz predpomnilnika …", flush=True)
    try:
        model_data = _load_cache_model(cache_path)
        print(f"   Model naložen  ({model_data['n_vzorcev']} vzorcev, "
              f"{args.n_trees} dreves) ⚡", flush=True)
    except Exception as e:
        print(f"   ⚠  Predpomnilnik napačen ({e}), ponovni trening …", flush=True)
        use_cache = False

if not use_cache:
    print(f"✓  Berem: {csv_path}", flush=True)
    meta, model_data = load_and_train(csv_path, args.n_trees, args.depth, args.seed)
    print(f"✓  Model naučen ({model_data['n_vzorcev']} vzorcev, {args.n_trees} dreves)",
          flush=True)
    try:
        _save_cache(cache_path, meta, model_data)
        print(f"✓  Model shranjen: {os.path.basename(cache_path)}", flush=True)
    except Exception as e:
        print(f"   ⚠  Model ni bil shranjen: {e}", flush=True)

# Razpakuj model
rf         = model_data["rf"]
f_means    = model_data["f_means"]
f_stds     = model_data["f_stds"]
imp        = model_data["imp"]
vrsta_enc  = model_data["vrsta_enc"]
obcina_enc = model_data["obcina_enc"]
ob_sorted  = model_data["ob_sorted"]
vc         = model_data["vc"]
raw_rows   = model_data["raw_rows"]
n_vzorcev  = model_data["n_vzorcev"]

# ── Samo-trening: izhod brez napovedi ─────────────────────────────────────────
if args.samo_trening:
    print()
    print(f"TRENING_OK=1")
    print(f"VZORCI={n_vzorcev}")
    print(f"MODEL_POT={cache_path}")
    sys.exit(0)


# ──────────────────────────────────────────────────────────────────────────────
#  NAPOVED
# ──────────────────────────────────────────────────────────────────────────────
vrsta_vhod = args.vrsta  or sorted(vc, key=lambda k: -vc[k])[0]
kraj_vhod  = args.kraj   or ob_sorted[len(ob_sorted) // 2]
povrsina   = args.povrsina   if args.povrsina   else imp["VelikostM2"]
zemljisce  = args.zemljisce  if args.zemljisce  else imp["ZemljisteM2"]
leto       = args.leto       if args.leto       else imp["LetoGradnje"]
sobe       = args.sobe       if args.sobe       else imp["StSob"]
enr_in     = _ENERGY.get(str(args.energija or "").strip().upper(),
                          imp["EnergetRazred"])

x_raw = [
    povrsina,
    max(0.0, zemljisce),
    leto if leto > 1_900 else imp["LetoGradnje"],
    sobe,
    float(enr_in),
    float(vrsta_enc.get(vrsta_vhod, vrsta_enc.get(list(vrsta_enc)[0], 0))),
    float(obcina_enc.get(kraj_vhod, len(obcina_enc) // 2)),
]
x_std = [(x_raw[j] - f_means[j]) / f_stds[j] for j in range(len(x_raw))]

tree_preds = rf.predict_all(x_std)
napoved    = statistics.mean(tree_preds)
std_pred   = statistics.stdev(tree_preds) if len(tree_preds) > 1 else 0.0
ci_min     = max(0.0, napoved - 1.64 * std_pred)
ci_max     = napoved + 1.64 * std_pred

# ── Podobni oglasi ─────────────────────────────────────────────────────────────
def _similar(vrsta=None, kraj=None, min_n=3):
    if vrsta and kraj:
        r = [x for x in raw_rows if x["VrstaObjekta"] == vrsta and x["Obcina"] == kraj]
        if len(r) >= min_n: return r, "ista vrsta + kraj"
    if vrsta:
        r = [x for x in raw_rows if x["VrstaObjekta"] == vrsta]
        if len(r) >= min_n: return r, "ista vrsta"
    if kraj:
        r = [x for x in raw_rows if x["Obcina"] == kraj]
        if len(r) >= min_n: return r, "isti kraj"
    return raw_rows, "vsi oglasi"

podobni, sim_opis = _similar(vrsta_vhod if args.vrsta else None,
                              kraj_vhod  if args.kraj  else None)
sim_cene   = [x["Cena"] for x in podobni]
sim_povp   = statistics.mean(sim_cene)
sim_med    = statistics.median(sim_cene)
napaka_pct = abs(napoved - sim_med) / sim_med * 100 if sim_med else 0.0

# ── Izpis (KEY=VALUE format za GUI) ───────────────────────────────────────────
print()
print(f"VZORCI={n_vzorcev}")
print(f"NAPOVEDANA={napoved:.0f}")
print(f"CI_MIN={ci_min:.0f}")
print(f"CI_MAX={ci_max:.0f}")
print(f"STD={std_pred:.0f}")
print(f"PODOBNI_N={len(podobni)}")
print(f"PODOBNI_FILTER={sim_opis}")
print(f"PODOBNI_MIN={min(sim_cene):.0f}")
print(f"PODOBNI_MAX={max(sim_cene):.0f}")
print(f"PODOBNI_POVP={sim_povp:.0f}")
print(f"PODOBNI_MED={sim_med:.0f}")
print(f"NAPAKA_PCT={napaka_pct:.1f}")
print(f"VRSTA={vrsta_vhod}")
print(f"KRAJ={kraj_vhod}")
print(f"POVRSINA={povrsina:.0f}")
print(f"SOBE={sobe:.0f}")
print(f"LETO={leto:.0f}")
