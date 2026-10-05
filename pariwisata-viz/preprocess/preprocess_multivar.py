"""Bagian 10 - data multivariat provinsi (PCA, klaster, pencilan) -> web/data/multivar.json
Jalankan dari folder proyek:  python preprocess/preprocess_multivar.py [folder_xlsx]   (default: data/raw)
Butuh: openpyxl, numpy (tanpa scikit-learn). Berdiri sendiri; tidak perlu menjalankan preprocess.py dulu."""
import json, sys, pathlib, re
import numpy as np
from openpyxl import load_workbook

RAW = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "data/raw")
OUT = pathlib.Path("web/data"); OUT.mkdir(parents=True, exist_ok=True)
TAHUN = 2025
BULAN = "Januari Februari Maret April Mei Juni Juli Agustus September Oktober November Desember".split()
F_WISNUS = "Jumlah_perjalanan_Wisnus.xlsx"
F_TPK = "TPK_dan_rata-rata_lama_menginap.xlsx"
F_MAKSUD = "Persentase_Perjalanan_Wisatawan_Nusantara_2025.xlsx"
F_PEND = "jumlah_penduduk_per_provinsi.xlsx"
F_AKOMODASI = "jumlah_akomodasi_per_provinsi.xlsx"
F_BELANJA = "rata_rata_pengeluaran_per_perjalanan_wisatawan_nusantara_2025.xlsx"

# ---------------------------------------------------------------- pembaca berkas
def find(fname):
    if not RAW.exists():
        raise SystemExit(f"Folder {RAW.resolve()} tidak ada. Buat folder itu lalu salin file xlsx ke dalamnya.")
    p = RAW / fname
    if p.exists(): return p
    cocok = sorted(RAW.glob(pathlib.Path(fname).stem + "*.xlsx"))
    if cocok: return cocok[0]
    raise SystemExit(f"File {fname} tidak ditemukan.\nIsi {RAW.resolve()}:\n  " + "\n  ".join(x.name for x in RAW.iterdir()))

def sheet(fname, i=0):
    ws = load_workbook(find(fname), read_only=True, data_only=True).worksheets[i]
    return [list(r) for r in ws.iter_rows(values_only=True)]

num = lambda x: float(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else None   # '-', 'NA', kosong -> None

def canon(s):
    """Kunci nama provinsi yang seragam antarberkas: 'KEP. RIAU' = 'Kepulauan Riau', 'DI YOGYAKARTA' = 'DI Yogyakarta'."""
    k = re.sub(r"[^a-z]", "", str(s).lower())
    return "kepulauan" + k[3:] if k.startswith("kep") and not k.startswith("kepulauan") else k

def tabel(rows, start, ncol):
    """Baris data mulai indeks `start` sampai sel pertama kosong / 'Catatan'. -> {kunci: (nama asli, [nilai...])}"""
    d = {}
    for r in rows[start:]:
        if r[0] is None or str(r[0]).startswith("Catatan"): break
        if str(r[0]).strip().upper() == "INDONESIA": continue
        d[canon(r[0])] = (str(r[0]).strip(), [num(x) for x in r[1:1 + ncol]])
    return d

# ---------------------------------------------------------------- 1. daftar provinsi (acuan = tabel penduduk, 38 provinsi)
pend_rows = sheet(F_PEND)
pend = {k: (n, v[0]) for k, (n, v) in tabel(pend_rows, 3, 1).items()}
prov = list(pend)                                    # urutan baris BPS (barat -> timur)
print("Provinsi (penduduk):", len(prov))

# ---------------------------------------------------------------- 2. wisnus 2025: jumlah perjalanan asal & tujuan
def wisnus_tahunan(idx):
    r = sheet(F_WISNUS, idx)
    yr, cols, tahunan = None, [], None
    for j in range(1, len(r[1])):
        yr = r[1][j] or yr
        if int(yr) == TAHUN:
            if r[2][j] in BULAN: cols.append(j)
            elif r[2][j] == "Tahunan": tahunan = j
    assert len(cols) == 12, f"Data wisnus {TAHUN} tidak lengkap 12 bulan ({len(cols)})"
    out = {}
    for row in r[3:]:
        if row[0] is None or str(row[0]).startswith("Catatan"): break
        if str(row[0]).strip().upper() == "INDONESIA": continue
        s = sum(num(row[j]) or 0 for j in cols)
        if tahunan is not None and num(row[tahunan]) and abs(s - row[tahunan]) > 1:
            print(f"  PERINGATAN {row[0]}: jumlah 12 bulan {s:.0f} != kolom Tahunan {row[tahunan]}")
        out[canon(row[0])] = s
    return out
asal, tuj = wisnus_tahunan(0), wisnus_tahunan(1)

# ---------------------------------------------------------------- 3. TPK & lama menginap 2025
t1 = sheet(F_TPK, 0); yrs = [int(y) for y in t1[1][1:] if y]
tpk = {k: v[yrs.index(TAHUN)] for k, (_, v) in tabel(t1, 2, len(yrs)).items()}
t2 = sheet(F_TPK, 1); g, jum = None, None
for j in range(1, len(t2[2])):
    g = t2[1][j] or g
    if g == "Jumlah" and int(t2[2][j]) == TAHUN: jum = j
assert jum, "kolom 'Jumlah' tahun 2025 tidak ditemukan pada lembar lama menginap"
inap = {canon(r[0]): num(r[jum]) for r in t2[3:] if r[0] and not str(r[0]).startswith("Catatan") and str(r[0]).upper() != "INDONESIA"}
# kolom tambahan hotel non-bintang (2025): TPK pada lembar 1, lama menginap tamu Indonesia pada lembar 2
jt = next(j for j, c in enumerate(t1[0]) if c and "nonbintang" in str(c).lower().replace(" ", ""))
assert int(t1[1][jt]) == TAHUN, "kolom TPK non-bintang bukan tahun 2025"
tpk_nb = {canon(r[0]): num(r[jt]) for r in t1[2:] if r[0] and not str(r[0]).startswith("Catatan") and str(r[0]).upper() != "INDONESIA"}
ji = next(j for j, c in enumerate(t2[1]) if c and "non bintang" in str(c).lower())
assert int(t2[2][ji]) == TAHUN, "kolom lama menginap non-bintang bukan tahun 2025"
inap_nb = {canon(r[0]): num(r[ji]) for r in t2[3:] if r[0] and not str(r[0]).startswith("Catatan") and str(r[0]).upper() != "INDONESIA"}

# ---------------------------------------------------------------- 4. akomodasi 2025 (hotel bintang & non-bintang) dan pengeluaran
def akomodasi(idx):
    r = sheet(F_AKOMODASI, idx)
    assert [str(c).strip() for c in r[1][1:4]] == ["Akomodasi", "Kamar", "Tempat Tidur"], "kolom lembar akomodasi berubah"
    assert all(int(y) == TAHUN for y in r[2][1:4]), f"tahun akomodasi bukan {TAHUN}"
    return tabel(r, 3, 3)                              # {kunci: (nama, [akomodasi, kamar, tempat tidur])}
akB, akN = akomodasi(0), akomodasi(1)                  # lembar 1 = hotel bintang, lembar 2 = hotel non-bintang

def hdr_table(fname):
    r = sheet(fname); h = [str(c).strip() for c in r[2]]
    return h, tabel(r, 3, len(h) - 1)
hB, belanja = hdr_table(F_BELANJA)
rm = sheet(F_MAKSUD); hi = next(i for i, r in enumerate(rm) if r[0] == "Provinsi Asal")
maksud = tabel(rm, hi + 1, len(rm[hi]) - 1)             # hanya untuk nama tampil provinsi (Title Case)
ix = lambda h, kunci: next(i for i, c in enumerate(h[1:]) if kunci.lower() in c.lower())   # indeks kolom data (tanpa kolom nama)
iTot = ix(hB, "Total")

# ---------------------------------------------------------------- 5. matriks variabel mentah (hanya variabel numerik)
# kunci, label lengkap, label ringkas ('|' = pemenggalan baris; maks. 12 huruf per baris), satuan, transformasi ('log' atau None), desimal
VARS = [
    ("asal",    "Perjalanan asal per 1.000 penduduk",          "Perjalanan|asal per|1.000 pddk",      "",      None,  0),
    ("tuj",     "Perjalanan tujuan per 1.000 penduduk",        "Perjalanan|tujuan per|1.000 pddk",    "",      None,  0),
    ("belanja", "Pengeluaran per perjalanan",                  "Pengeluaran|per|perjalanan",          "rb Rp", "log", 0),
    ("kmrB",    "Kamar hotel bintang per 100 rb penduduk",     "Kamar hotel|bintang per|100 rb pddk", "",      None,  0),
    ("kmrN",    "Kamar hotel non-bintang per 100 rb penduduk", "Kamar non-|bintang per|100 rb pddk",  "",      None,  0),
    ("ukur",    "Rata-rata kamar per hotel bintang",           "Kamar per|hotel|bintang",             "",      None,  0),
    ("tpk",     "TPK hotel bintang",                           "TPK hotel|bintang",                   "%",     None,  1),
    ("tpkN",    "TPK hotel non-bintang dan akomodasi lain",    "TPK hotel|non-bintang",               "%",     None,  1),
    ("inap",    "Lama menginap di hotel bintang",              "Lama|menginap|bintang",               "malam", None,  2),
    ("inapN",   "Lama menginap tamu Indonesia, hotel non-bintang", "Lama|menginap|non-bintang",       "malam", None,  2),
]
KEY = [v[0] for v in VARS]
div = lambda a, b: None if a is None or not b else a / b
raw, imp = {}, {}
for k in prov:
    pop = pend[k][1]; (hb, kb), kn = (akB[k][1][0], akB[k][1][1]), akN[k][1][1]     # akomodasi & kamar hotel bintang; kamar non-bintang
    raw[k] = [asal[k] / pop * 1000, tuj[k] / pop * 1000, belanja[k][1][iTot],
              None if kb is None else kb / pop * 1e5, None if kn is None else kn / pop * 1e5, div(kb, hb),
              tpk[k], tpk_nb[k], inap[k], inap_nb[k]]
M = np.array([[np.nan if x is None else x for x in raw[k]] for k in prov], float)

# Imputasi sel kosong: rata-rata provinsi lain di kelompok yang sama (Papua*); cadangan = median seluruh provinsi
kosong = np.argwhere(np.isnan(M))
for i, j in kosong:
    kel = [m for m, k in enumerate(prov) if m != i and pend[k][0].upper().startswith("PAPUA") == pend[prov[i]][0].upper().startswith("PAPUA")]
    v = M[kel, j]; v = v[~np.isnan(v)]
    M[i, j] = v.mean() if pend[prov[i]][0].upper().startswith("PAPUA") and len(v) else np.nanmedian(M[:, j])
    imp.setdefault(prov[i], []).append(int(j))
print("Sel diimputasi:", {pend[k][0]: [KEY[j] for j in js] for k, js in imp.items()} or "tidak ada")

# ---------------------------------------------------------------- 6. transformasi, z-score, PCA
X = np.column_stack([np.log10(M[:, j]) if VARS[j][4] == "log" else M[:, j] for j in range(len(VARS))])
Z = (X - X.mean(0)) / X.std(0, ddof=1)
n, p = Z.shape
U, S, Vt = np.linalg.svd(Z, full_matrices=False)
eig = S ** 2 / (n - 1); ratio = eig / eig.sum()
for c in range(Vt.shape[0]):                                   # tanda komponen: beban terbesar dibuat positif (agar stabil)
    if Vt[c, np.argmax(np.abs(Vt[c]))] < 0: Vt[c] *= -1; U[:, c] *= -1
scores = U * S                                                  # skor PC (n x p)
load = Vt.T * np.sqrt(eig)                                      # korelasi variabel dengan PC (p x p)
print("Varians dijelaskan PC1-4:", [f"{r:.1%}" for r in ratio[:4]], " kumulatif PC1-3: ", f"{ratio[:3].sum():.1%}")
for c in range(3):
    o = np.argsort(-np.abs(load[:, c]))[:4]
    print(f"  PC{c + 1}:", ", ".join(f"{KEY[j]} {load[j, c]:+.2f}" for j in o))

# ---------------------------------------------------------------- 7. k-means (numpy) + silhouette
def kmeans(A, k, rng, n_init=60):
    best = None
    for _ in range(n_init):
        C = [A[rng.integers(len(A))]]
        for _ in range(k - 1):                                  # k-means++
            d2 = np.min([((A - c) ** 2).sum(1) for c in C], axis=0)
            C.append(A[rng.choice(len(A), p=d2 / d2.sum())])
        C = np.array(C)
        for _ in range(100):
            lab = ((A[:, None, :] - C[None]) ** 2).sum(2).argmin(1)
            C2 = np.array([A[lab == c].mean(0) if (lab == c).any() else C[c] for c in range(k)])
            if np.allclose(C, C2): break
            C = C2
        inertia = ((A - C[lab]) ** 2).sum()
        if best is None or inertia < best[0]: best = (inertia, lab, C)
    return best[1], best[2]

def silhouette(A, lab):
    D = np.sqrt(((A[:, None, :] - A[None]) ** 2).sum(2)); s = []
    for i in range(len(A)):
        same = (lab == lab[i]); same[i] = False
        if not same.any(): s.append(0); continue
        a = D[i, same].mean(); b = min(D[i, lab == c].mean() for c in set(lab) if c != lab[i])
        s.append((b - a) / max(a, b))
    return float(np.mean(s))

rng = np.random.default_rng(42)
hasil = {k: kmeans(Z, k, rng) for k in range(2, 7)}
sil = {k: silhouette(Z, hasil[k][0]) for k in hasil}
K = max((3, 4, 5), key=lambda k: sil[k])                        # k terbaik di rentang 3-5 (k=2 terlalu kasar untuk dibahas)
print("Silhouette:", {k: round(v, 3) for k, v in sil.items()}, "-> K =", K)
lab, C = hasil[K]
urut = np.argsort([scores[lab == c, 0].mean() for c in range(K)])     # klaster diberi nomor menurut rata-rata PC1
peta = {int(c): i for i, c in enumerate(urut)}
lab = np.array([peta[int(c)] for c in lab]); C = C[urut]

# ---------------------------------------------------------------- 8. pencilan: jarak ke pusat (ruang z penuh), pagar Tukey
d = np.sqrt((Z ** 2).sum(1))
q1, q3 = np.percentile(d, [25, 75]); pagar = q3 + 1.5 * (q3 - q1)
pencilan = d > pagar
print(f"Pencilan (jarak > {pagar:.2f}):", [pend[prov[i]][0] for i in np.where(pencilan)[0]])
for c in range(K):
    top = np.argsort(-np.abs(C[c]))[:3]
    print(f"  Klaster {chr(65 + c)} (n={int((lab == c).sum())}):", ", ".join(f"{KEY[j]} {C[c, j]:+.2f}" for j in top),
          "|", ", ".join(pend[prov[i]][0] for i in np.where(lab == c)[0]))

# ---------------------------------------------------------------- 9. tulis JSON
nama = {}                                                       # nama tampil Title Case dari tabel moda
for k in prov: nama[k] = maksud[k][0] if k in maksud else pend[k][0].title()
r3, r4 = (lambda a: [round(float(x), 3) for x in a]), (lambda a: [round(float(x), 4) for x in a])
rows = []
for i, k in enumerate(prov):
    zi = Z[i]; t = int(np.argmax(np.abs(zi)))
    rows.append({"p": pend[k][0], "n": nama[k], "pop": int(pend[k][1]), "raw": r3(M[i]), "x": r4(X[i]), "z": r3(zi),
                 "pc": r3(scores[i, :3]), "cl": int(lab[i]), "out": bool(pencilan[i]), "d": round(float(d[i]), 3),
                 "top": [t, round(float(zi[t]), 2)], "imp": imp.get(k, [])})
dump = {"year": TAHUN, "vars": [dict(zip(("k", "l", "s", "u", "tf", "dec"), v)) for v in VARS], "prov": rows,
        "pca": {"ratio": r4(ratio), "load": [r3(load[j, :3]) for j in range(p)]},
        "km": {"k": K, "sil": {str(a): round(b, 3) for a, b in sil.items()}, "centers": [r3(c) for c in C],
               "size": [int((lab == c).sum()) for c in range(K)], "fence": round(float(pagar), 3)}}
(OUT / "multivar.json").write_text(json.dumps(dump, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
print(f"Selesai -> {OUT / 'multivar.json'}  ({n} provinsi x {p} variabel)")
