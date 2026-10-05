"""Ubah 3 berkas xlsx BPS menjadi JSON kecil untuk web.
Jalankan dari folder proyek:  python preprocess/preprocess.py [folder_xlsx]   (default: data/raw)"""
import json, sys, pathlib
from openpyxl import load_workbook

RAW = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "data/raw")
OUT = pathlib.Path("web/data"); OUT.mkdir(parents=True, exist_ok=True)
BULAN = {b: i + 1 for i, b in enumerate(
    "Januari Februari Maret April Mei Juni Juli Agustus September Oktober November Desember".split())}
F_WISMAN = "Kunjungan_Wisman_menurut_kebangsaan_dan_pintu_masuk.xlsx"
F_WISNUS = "Jumlah_perjalanan_Wisnus.xlsx"
F_TPK = "TPK_dan_rata-rata_lama_menginap.xlsx"
F_TREN = "Kunjungan_Wisatawan_Mancanegara_2015-2025.xlsx"
F_MAKSUD = "Persentase_Perjalanan_Wisatawan_Nusantara_2025.xlsx"
F_SUNBURST = "Sunburst_Jumlah_Perjalanan_Mancanegara_per_Negara_dan_Pintu_Masuk.xlsx"
F_GENDER = "persentase_perjalanan_wisatawan_asal_jenis_kelamin_2025.xlsx"

def find(fname):
    if not RAW.exists():
        raise SystemExit(f"Folder {RAW.resolve()} tidak ada. Buat folder itu lalu salin file xlsx ke dalamnya.")
    p = RAW / fname
    if p.exists(): return p
    cocok = sorted(RAW.glob(pathlib.Path(fname).stem + "*.xlsx"))     # menerima nama seperti '... (1).xlsx'
    if cocok: return cocok[0]
    raise SystemExit(f"File {fname} tidak ditemukan.\nIsi {RAW.resolve()}:\n  " + "\n  ".join(x.name for x in RAW.iterdir()))

def sheet(fname, i):
    ws = load_workbook(find(fname), read_only=True, data_only=True).worksheets[i]
    return [list(r) for r in ws.iter_rows(values_only=True)]
    
num = lambda x: x if isinstance(x, (int, float)) else None   # '-' / kosong -> None

def monthly(rows):
    """Baris 2 = tahun (sel digabung), baris 3 = bulan; kolom 'Tahunan' dibuang."""
    year, cols, months = None, [], []
    for j in range(1, len(rows[1])):
        year = rows[1][j] or year
        if rows[2][j] in BULAN:
            cols.append(j); months.append(f"{year}-{BULAN[rows[2][j]]:02d}")
    body = []
    for r in rows[3:]:
        if r[0] is None: break                      # akhir tabel (sebelum 'Catatan')
        body.append((str(r[0]).strip(), [num(r[j]) for j in cols]))
    return months, body

def check(label, parts, whole):
    d = [abs(sum(p[i] or 0 for p in parts) - whole[i]) / whole[i] for i in range(len(whole)) if whole[i]]
    print(f"  cek {label}: selisih maks {max(d):.3%}" if d else f"  cek {label}: tanpa data")

def dump(name, obj):
    (OUT / f"{name}.json").write_text(json.dumps(obj, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

# ---- 1. Wisman menurut pintu masuk: hierarki Jenis pintu -> pintu
months, body = monthly(sheet(F_WISMAN, 0))
JENIS = {"A. Pintu Udara": "Pintu Udara", "B. Pintu Laut": "Pintu Laut", "C. Pintu Darat": "Pintu Darat"}
rows, agg, grp = [], {}, None
for name, v in body:
    if name in JENIS: grp = JENIS[name]; agg[grp] = v
    elif name == "Jumlah (A+B+C)": agg[name] = v; grp = None
    elif name == "Total": agg["Total"] = v
    elif name == "Lainnya": rows.append({"path": ["Lainnya", "Lainnya"], "v": v})
    elif grp: rows.append({"path": [grp, name], "v": v})
print("Pintu masuk:", len(rows), "pintu,", len(months), "bulan", months[0], "-", months[-1])
for g in JENIS.values(): check(g, [r["v"] for r in rows if r["path"][0] == g], agg[g])
dump("pintu", {"months": months, "agg": agg, "rows": rows})

# ---- 2. Wisman menurut paspor: hierarki Kawasan -> negara
months2, body = monthly(sheet(F_WISMAN, 1))
TOT = {"A S E A N": "ASEAN", "TOTAL ASIA (Excl.ASEAN)": "Asia (di luar ASEAN)", "TOTAL MIDDLE EAST": "Timur Tengah",
       "TOTAL EUROPE": "Eropa", "TOTAL AMERICA": "Amerika", "TOTAL OCEANIA": "Oseania", "TOTAL AFRICA": "Afrika"}
# Baris 'Other ...' = agregat negara kecil; negara di bawahnya sudah termasuk di dalamnya -> dilewati agar tidak dobel
HEAD = {"Other Asia", "Other Middle East", "Other West Europe", "Other East Europe", "Central America",
        "South America", "Other America", "Other Oceania", "Other Africa"}
paspor, cur, skip, tots = [], [], False, {}
for name, v in body:
    if name in TOT:
        paspor += [{"path": [TOT[name], n], "v": vv} for n, vv in cur]
        tots[TOT[name]] = v; cur, skip = [], False
    elif name == "GRAND TOTAL": tots["Total"] = v
    elif name in HEAD: cur.append((f"Lainnya ({name})", v)); skip = True
    elif not skip: cur.append((name, v))
print("Paspor:", len(paspor), "negara/kelompok")
for g, tv in tots.items():
    if g != "Total": check(g, [r["v"] for r in paspor if r["path"][0] == g], tv)
check("Grand total", [r["v"] for r in paspor], tots["Total"])
dump("paspor", {"months": months2, "agg": tots, "rows": paspor})

# ---- 3. Wisnus: provinsi asal & tujuan (bulanan)
m3, asal = monthly(sheet(F_WISNUS, 0)); _, tuj = monthly(sheet(F_WISNUS, 1))
A, T = dict(asal), dict(tuj)
prov = [n for n, _ in asal if n != "INDONESIA"]
print("Wisnus:", len(prov), "provinsi")
check("asal", [A[p] for p in prov], A["INDONESIA"]); check("tujuan", [T[p] for p in prov], T["INDONESIA"])
dump("wisnus", {"months": m3, "rows": [{"p": p, "asal": A[p], "tujuan": T[p]} for p in prov]})

# ---- 4. TPK & lama menginap (tahunan 2021-2025)
r1 = sheet(F_TPK, 0); years = [int(y) for y in r1[1][1:] if y]
TPK = {}
for r in r1[2:]:
    if r[0] is None: break
    TPK[str(r[0]).strip()] = [num(x) for x in r[1:1 + len(years)]]
r2 = sheet(F_TPK, 1); g, keys = None, []
for j in range(1, len(r2[2])):
    g = r2[1][j] or g; keys.append((g, int(r2[2][j])))
GKEY = {"Tamu Asing": "asing", "Tamu Indonesia": "indo", "Jumlah": "total"}
INAP = {}
for r in r2[3:]:
    if r[0] is None: break
    d = {k: [None] * len(years) for k in GKEY.values()}
    for j, (gg, yy) in enumerate(keys, start=1): d[GKEY[gg]][years.index(yy)] = num(r[j])
    INAP[str(r[0]).strip()] = d
rows = [{"p": p, "tpk": TPK[p], **INAP.get(p, {})} for p in TPK if p != "INDONESIA"]
print("TPK/menginap:", len(rows), "provinsi, tahun", years)
dump("tpk", {"years": years, "rows": rows})

# ---- 5. Tren wisman 2015-2025 (total, bulanan) -> grafik tren pembuka
r5 = sheet(F_TREN, 0)
hdr = next(i for i, r in enumerate(r5) if r[0] == "Tahun")
m5, v5, annual, bad = [], [], {}, []
for r in r5[hdr + 1:]:
    if not isinstance(r[0], int): continue
    for k in range(12):
        m5.append(f"{r[0]}-{k + 1:02d}"); v5.append(num(r[1 + k]))
    annual[str(r[0])] = r[13]
    if sum(x or 0 for x in v5[-12:]) != r[13]: bad.append(r[0])
print("Tren wisman:", len(annual), "tahun;", "jumlah bulan = jumlah tahunan di file" if not bad else f"PERINGATAN jumlah tidak cocok: {bad}")
dump("wisman_tahunan", {"months": m5, "v": v5, "annual": annual})

# ---- 6. Asal wisman per negara, tahunan 2021-2025 -> peta aliran
# nama di tabel BPS: (nama tampil, bujur, lintang). Titik = ibu kota negara (aproksimasi untuk menggambar arah)
COORD = {
 "Malaysia": ("Malaysia", 101.687, 3.139), "Australia": ("Australia", 149.128, -35.281),
 "Singapore": ("Singapura", 103.820, 1.352), "Republik Rakyat Tiongkok": ("Tiongkok", 116.407, 39.904),
 "Timor Leste": ("Timor Leste", 125.573, -8.557), "India": ("India", 77.209, 28.614),
 "South Korea": ("Korea Selatan", 126.978, 37.566), "United Kingdom": ("Inggris Raya", -0.128, 51.507),
 "United States of America": ("Amerika Serikat", -77.037, 38.907), "Japan": ("Jepang", 139.692, 35.690),
 "France": ("Prancis", 2.352, 48.857), "Germany": ("Jerman", 13.405, 52.520),
 "Netherlands": ("Belanda", 4.904, 52.368), "Philippines": ("Filipina", 120.984, 14.600),
 "R u s i a": ("Rusia", 37.617, 55.756), "Taiwan": ("Taiwan", 121.565, 25.033),
 "New Zealand": ("Selandia Baru", 174.776, -41.287), "Saudi Arabia": ("Arab Saudi", 46.675, 24.714),
 "Papua New Guinea": ("Papua Nugini", 147.180, -9.443), "I t a l y": ("Italia", 12.496, 41.903),
 "Spain": ("Spanyol", -3.704, 40.417), "Hongkong, RRT": ("Hong Kong", 114.169, 22.319),
 "Thailand": ("Thailand", 100.502, 13.756), "Canada": ("Kanada", -75.697, 45.421),
 "Vietnam": ("Vietnam", 105.834, 21.028), "Polandia": ("Polandia", 21.012, 52.230),
 "Turki": ("Turki", 32.860, 39.933), "Switzerland": ("Swiss", 7.447, 46.948),
 "Myanmar/Burma": ("Myanmar", 96.078, 19.763), "Belgium": ("Belgia", 4.352, 50.850),
 "Denmark": ("Denmark", 12.568, 55.676), "Portugal": ("Portugal", -9.139, 38.722),
 "Irlandia (Ireland)": ("Irlandia", -6.260, 53.350), "Sweden": ("Swedia", 18.069, 59.329),
 "South Africa": ("Afrika Selatan", 28.188, -25.746), "Pakistan": ("Pakistan", 73.048, 33.684),
 "Brazilia": ("Brasil", -47.882, -15.794), "Austria": ("Austria", 16.373, 48.208),
 "Brunei Darussalam": ("Brunei", 114.948, 4.940), "Ukraine": ("Ukraina", 30.523, 50.450),
 "Ceko": ("Ceko", 14.420, 50.088), "Romania": ("Rumania", 26.103, 44.427),
 "Kazakhstan": ("Kazakhstan", 71.447, 51.160), "Norway": ("Norwegia", 10.752, 59.914),
 "Egypt": ("Mesir", 31.236, 30.044), "Hongaria": ("Hongaria", 19.040, 47.498),
 "Bangladesh": ("Bangladesh", 90.413, 23.811), "Maroko": ("Maroko", -6.849, 34.021),
 "Mexiko": ("Meksiko", -99.133, 19.433), "Oman": ("Oman", 58.405, 23.588),
 "Kolombia": ("Kolombia", -74.072, 4.711), "Nepal": ("Nepal", 85.324, 27.717),
 "Finland": ("Finlandia", 24.938, 60.170), "Argentina": ("Argentina", -58.382, -34.604),
 "Yunani (Greece)": ("Yunani", 23.727, 37.984),
}
YEARS6 = [2021, 2022, 2023, 2024, 2025]
yearly = lambda v: [sum(x or 0 for x, m in zip(v, months2) if m.startswith(str(y))) for y in YEARS6]
rows6, total6, unmapped6 = [], [0] * len(YEARS6), [0] * len(YEARS6)
for r in paspor:
    yv = yearly(r["v"]); total6 = [a + b for a, b in zip(total6, yv)]
    raw = r["path"][1]
    if raw in COORD:
        nm, lon, lat = COORD[raw]; rows6.append({"n": nm, "r": r["path"][0], "c": [lon, lat], "v": yv})
    else: unmapped6 = [a + b for a, b in zip(unmapped6, yv)]
print(f"Asal wisman: {len(rows6)} negara dipetakan; tidak dipetakan {sum(unmapped6) / sum(total6):.1%} (kelompok 'Lainnya', negara kecil, paspor Indonesia)")
dump("asal", {"years": YEARS6, "rows": rows6, "total": total6, "unmapped": unmapped6})

# ---- 7. Maksud perjalanan wisnus 2025 (persen) -> sunburst
r7 = sheet(F_MAKSUD, 0)
hdr7 = next(i for i, r in enumerate(r7) if r[0] == "Provinsi Asal")
kat = [str(c).strip() for c in r7[hdr7][1:12]]              # 11 maksud perjalanan (kolom 'Total' dibuang)
rows7 = []
for r in r7[hdr7 + 1:]:
    if r[0] is None or str(r[0]).startswith("Catatan"): break
    nama = "Indonesia" if str(r[0]).strip() == "INDONESIA" else str(r[0]).strip()
    rows7.append({"p": nama, "v": [num(x) for x in r[1:12]]})    # 'NA' dan '-' -> None
rows7.sort(key=lambda d: d["p"] != "Indonesia")                  # Indonesia di urutan pertama
print("Maksud perjalanan:", len(rows7), "wilayah,", len(kat), "kategori;",
      "ada NA/- pada:", [d["p"] for d in rows7 if None in d["v"]])
dump("maksud", {"kategori": kat, "rows": rows7})

# ---- 8. Wisman menurut negara x pintu masuk -> sunburst pintu masuk
r8 = sheet(F_SUNBURST, 0)
H8 = [str(c).strip() for c in r8[0]]
assert H8[30:34] == ["Total Pintu Udara", "Total Pintu Laut", "Total Pintu Darat", "Total Pintu Masuk Utama"], "urutan kolom berubah"
# (jenis, kolom awal, kolom akhir + 1, kolom total di Excel); kolom terakhir tiap jenis = 'Lainnya'
JENIS8 = [("Udara", 1, 17, 30), ("Laut", 17, 24, 31), ("Darat", 24, 30, 32)]
LAIN8 = {"Udara": "Lainnya (bandara)", "Laut": "Lainnya (pelabuhan)", "Darat": "Lainnya (pos darat)"}
pintu8 = [{"n": H8[j], "s": LAIN8[jn] if j == b - 1 else H8[j].split(",")[0].strip(), "j": jn}
          for jn, a, b, _ in JENIS8 for j in range(a, b)]
# Kawasan tiap negara: tabel urut per kawasan dan tiap kawasan diakhiri baris 'Lainnya' (nama di bawah)
KAWASAN8 = [("ASEAN", "Asean Lainnya"), ("Asia (di luar ASEAN)", "Asia Lainnya"), ("Timur Tengah", "Timur Tengah Lainnya"),
            ("Eropa", "Eropa Timur Lainnya"), ("Amerika", "Amerika Lainnya"), ("Oseania", "Oceania Lainnya"), ("Afrika", "Afrika Lainnya")]
negara8, koreksi8, kw = [], [], 0
for r in r8[1:]:
    if r[0] is None: break
    nama = str(r[0]).split("/")[0].strip()                          # 'Singapura/Singapore' -> 'Singapura'
    v = [num(x) or 0 for x in r[1:30]]
    for jn, a, b, ct in JENIS8:
        lain = max((r[ct] or 0) - sum(v[a - 1:b - 2]), 0)           # 'Lainnya' = total jenis pintu - pintu yang disebut
        if lain != v[b - 2]:
            koreksi8.append({"n": nama, "j": jn, "asal": v[b - 2], "baru": lain}); v[b - 2] = lain
    negara8.append({"n": nama, "r": KAWASAN8[kw][0], "v": v})
    if nama == KAWASAN8[kw][1]: kw += 1                            # baris penutup kawasan -> pindah kawasan berikutnya
print("Sunburst pintu:", len(negara8), "negara,", len(pintu8), "pintu;", len(koreksi8), "sel 'Lainnya' dihitung ulang dari total jenis pintu")
assert kw == len(KAWASAN8), "pembagian kawasan tidak cocok dengan urutan baris di Excel"
dump("pintu_negara", {"pintu": pintu8, "negara": negara8, "koreksi": koreksi8})

# ---- 9. Jenis kelamin wisnus per provinsi asal 2025 (persen) -> sankey
r9 = sheet(F_GENDER, 0)
hdr9 = next(i for i, r in enumerate(r9) if r[0] == "Provinsi Asal")
rows9 = []
for r in r9[hdr9 + 1:]:
    if r[0] is None: break
    nama = "Indonesia" if str(r[0]).strip() == "INDONESIA" else str(r[0]).strip()   # sama dengan bagian 7
    rows9.append({"p": nama, "L": num(r[1]), "P": num(r[2])})
sama = {d["p"] for d in rows9} == {d["p"] for d in rows7}
print("Gender:", len(rows9), "wilayah; nama provinsi", "cocok dengan tabel maksud" if sama else "TIDAK cocok dengan tabel maksud")
assert all(abs(d["L"] + d["P"] - 100) < 0.05 for d in rows9), "L+P tidak 100%"
dump("gender", {"rows": rows9})

print("Selesai -> web/data/")
