#!/usr/bin/env python3
"""
Costruisce data/cerchi.plain.json e rigenera index.html per la web app
"Cerchi in acciaio Magnetto - MAK".

    MAGNETTO_PIN='****' python3 data/build.py
    # oppure, senza variabile d'ambiente, lo script chiede il PIN (input nascosto)

Fonte dati:
  data/Catalogo Tecnico Magnetto Acciaio.xlsx  (primo foglio)
  "MW Master Katalog" di Magnetto Wheels: una riga per abbinamento vettura-cerchio.

Cosa fa:
  1. legge il foglio, scarta le righe vuote (separatori di marca) e quelle senza
     codice articolo;
  2. normalizza marca (suffisso "Nfz" -> veicoli commerciali), canale, ET,
     misure pneumatico, date di produzione, finitura;
  3. sceglie la nota in italiano; dove il catalogo ha solo il tedesco la traduce
     con NOTE_IT / le regole di traduci_nota() (l'originale resta nel dataset);
  4. cifra il dataset: chiave = PBKDF2-SHA256(PIN, salt, 200k) -> AES-GCM-256;
  5. inietta il blob cifrato nel template e scrive index.html.

I dati non vengono corretti: le incoerenze del catalogo (stesso articolo con pesi
o codici Alcar diversi) restano come sono, riga per riga. Vedi README.md.

Il PIN non viene mai stampato ne' salvato. Nel repo pubblico finisce solo
index.html con il blob cifrato. data/*.xlsx e data/*.plain.json sono in .gitignore.

Dipendenze: pip3 install openpyxl cryptography
"""

import base64
import csv
import getpass
import hashlib
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)

XLSX = os.path.join(HERE, "Catalogo Tecnico Magnetto Acciaio.xlsx")
TEMPLATE = os.path.join(REPO, "index.template.html")
INDEX = os.path.join(REPO, "index.html")
PLAIN_JSON = os.path.join(HERE, "cerchi.plain.json")
XLSX_KRONPRINZ = os.path.join(HERE, "Listino Acciaio Kronprinz.xlsx")
CSV_EQUIVALENZE = os.path.join(HERE, "equivalenze.csv")
# solo misure e finitura: le colonne dei prezzi A1/A2/A3 non vengono lette
XLSX_LISTINO_MAK = os.path.join(REPO, "..", "ebook", "download",
                                "prodotti-e-promozioni-attive__mak__ruote-in-ferro-accessori",
                                "LISTINI ACCIAIO(A1+A2+A3).xlsx")
RE_PORTATA = re.compile(r"Carico ruota|Portata|Max\s+[\d,.]+\s*Ton", re.I)
FINITURE_KRONPRINZ = {"SILVER": "Argento", "MATT BLACK": "Nero"}

PBKDF2_ITER = 200_000
PRIMA_RIGA_DATI = 4

MESI_DE = {"Januar": "gennaio", "Februar": "febbraio", "März": "marzo", "April": "aprile",
           "Mai": "maggio", "Juni": "giugno", "Juli": "luglio", "August": "agosto",
           "September": "settembre", "Oktober": "ottobre", "November": "novembre",
           "Dezember": "dicembre"}

FINITURE = {"schwarz": "Nero", "silber": "Argento"}

# misure pneumatico scritte male nel catalogo
PNEU_FIX = {"205//70R15": "205/70R15", "18565R15": "185/65R15"}

# note presenti solo in tedesco: traduzione redazionale (non di Magnetto)
NOTE_IT = {
    '15"nicht bei elektischer Parkbremse / Radlast 685 kg':
        '15" non con freno di stazionamento elettrico - Carico ruota 685 kg',
    "185/55R15 je nach Ausführung zulässig": "185/55R15 ammesso secondo la versione",
    "205/75R16 nur Frontantrieb bis 3.5t Radlast 1.050 kg":
        "205/75R16 solo trazione anteriore fino a 3,5 t - Carico ruota 1.050 kg",
    "Achslast max. 2500 kg": "Portata asse max kg 2.500",
    "Asführung 80 nur mit VA 235/55R19 und HA 255/50R19":
        "Versione 80 solo con 235/55R19 anteriore e 255/50R19 posteriore",
    "Frontantrieb": "Trazione anteriore",
    "Heckantrieb": "Trazione posteriore",
    "Nicht zulässig bei Fahrzeugen mit elektrischer Feststellbremse":
        "Non ammesso sui veicoli con freno di stazionamento elettrico",
    "Originalblende passt nicht": "Il copriruota originale non è compatibile",
    "SUV auf Fiat 500 Basis": "SUV su base Fiat 500",
    'Stepway je nach Ausführung nur 16" zulässig': 'Stepway: alcune versioni montano solo 16"',
    "Styling Rad": "Ruota stile",
    "Stylingrad": "Ruota stile",
    "Tschechische Produktion": "Produzione Repubblica Ceca",
    "Zwillingsbereifung": "Montaggio in gemello",
    "Zwillingsbereifung - RDKS tauglich": "Montaggio in gemello - compatibile TPMS",
    "Zwillingsbereifung auch für Vorderachse": "Montaggio in gemello, anche per asse anteriore",
    "abgelöst durch Fiat 500X": "Sostituito da Fiat 500X",
    'einige Ausführungen nur in 17" möglich': 'Alcune versioni montano solo 17"',
    'je nach Ausführung und ab NG*39 nur 17" zulässig':
        'Alcune versioni, e tutte da NG*39, montano solo 17"',
    "kein Verkauf in Deutschland": "Non in vendita in Germania",
    "nicht 4-motion": "Non 4Motion",
    "nicht AMG Ausführung, nicht für Sportbremse geeignet":
        "Non per versioni AMG, non adatto all'impianto frenante sportivo",
    "nicht Passat Hybrid": "Non Passat Hybrid",
    "nicht für Hybrid 115kW": "Non per Hybrid 115 kW",
    "nur Heckantrieb über 3.5t Radlast 1.200 kg":
        "Solo trazione posteriore oltre 3,5 t - Carico ruota 1.200 kg",
    "nur Hinterachse": "Solo asse posteriore",
    "nur Vorderachse": "Solo asse anteriore",
    "nur für kurzen Radstand": "Solo passo corto",
    "nur kurzer Radstand mit Sensoren optional": "Solo passo corto, sensori opzionali",
    "schwarz": "Ruota nera",
    "silber": "Ruota grigia",
    "zulässig bis 3.5t Gesamtgewicht": "Max 3,5 Ton",
}


def s(v):
    """Cella -> stringa ripulita dagli spazi, None se vuota."""
    if v is None:
        return None
    t = re.sub(r"\s+", " ", str(v)).strip()
    return t or None


def num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(v) if float(v).is_integer() else round(float(v), 2)
    return None


def traduci_nota(de):
    """Nota tedesca -> italiano; None se non c'e' una regola."""
    if de in NOTE_IT:
        return NOTE_IT[de]
    m = re.fullmatch(r"Radlast ([\d.]+) kg", de)
    if m:
        return f"Carico ruota {m.group(1)} kg"
    m = re.fullmatch(r"je nach Ausführung nur (.+) zulässig", de)
    if m:
        return "Alcune versioni montano solo " + m.group(1).replace(" oder ", " o ")
    return None


def norm_marca(v):
    """'Fiat Nfz' / 'Volkswagen Nfz.' -> ('Fiat', True)."""
    t = s(v)
    m = re.fullmatch(r"(.+?)\s+Nfz\.?", t)
    return (m.group(1), True) if m else (t, False)


def norm_canale(v):
    """'6 1/2J ' -> ('6 1/2 J', 6.5); '5,00 B' -> ('5 B', 5)."""
    t = s(v)
    m = re.fullmatch(r"(\d+)(?:\s+(1/2)|,00)?\s*([A-Z]+)", t or "")
    if not m:
        return t, None
    intero, mezzo, profilo = m.groups()
    label = intero + (" 1/2" if mezzo else "") + " " + profilo
    return label, int(intero) + (0.5 if mezzo else 0)


def norm_et(v):
    """ET numerico, oppure testo ('HMA 121,5' sulle ruote gemellate)."""
    n = num(v)
    if n is not None:
        return n, None
    return None, s(v)


def norm_data(v):
    """'05.10' -> '05/2010'."""
    t = s(v)
    m = re.fullmatch(r"(\d{2})\.(\d{2})", t or "")
    if not m:
        return t
    yy = int(m.group(2))
    return f"{m.group(1)}/{1900 + yy if yy > 60 else 2000 + yy}"


def norm_pneu(v):
    t = s(v)
    if not t:
        return None
    t = t.replace(" ", "")
    return PNEU_FIX.get(t, t)


def parse():
    try:
        from openpyxl import load_workbook
    except ImportError:
        sys.exit("Manca openpyxl: pip3 install openpyxl")
    ws = load_workbook(XLSX, data_only=True).worksheets[0]

    edizione = s(ws["C1"].value) or ""
    m = re.search(r"(\w+)\s+(\d{4})", edizione)
    if m:
        edizione = f"{MESI_DE.get(m.group(1), m.group(1))} {m.group(2)}"
    prezzi = s(ws["Y3"].value) or ""
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{2})", prezzi)
    prezzi = f"{m.group(1)}.{m.group(2)}.20{m.group(3)}" if m else ""

    righe = []
    stats = {"vuote": 0, "senza_codice": [], "note_non_tradotte": [], "note_tradotte": 0}
    for i, r in enumerate(ws.iter_rows(min_row=PRIMA_RIGA_DATI, max_col=27, values_only=True),
                          PRIMA_RIGA_DATI):
        if all(c is None for c in r):
            stats["vuote"] += 1
            continue
        if not s(r[0]) or not s(r[15]):
            stats["senza_codice"].append(i)
            continue

        marca, comm = norm_marca(r[0])
        canale, canale_n = norm_canale(r[9])
        et, et_txt = norm_et(r[22])

        de = s(r[3])
        de = None if de == "." else de
        nota = s(r[4])
        if not nota and de:
            nota = traduci_nota(de)
            if nota:
                stats["note_tradotte"] += 1
            else:
                stats["note_non_tradotte"].append(de)
                nota = de

        righe.append({
            "b": marca, "cm": 1 if comm else 0,
            "m": s(r[1]), "v": s(r[2]),
            "n": nota, "nd": de,
            "abe": s(r[5]), "pj": s(r[6]),
            "f": norm_data(r[7]), "t": norm_data(r[8]),
            "w": canale, "wn": canale_n, "d": num(r[10]),
            "ty": [p for p in (norm_pneu(x) for x in r[11:14]) if p],
            "mwd": s(r[14]), "c": s(r[15]), "ean": s(r[16]), "alc": s(r[17]),
            "nw": 1 if (s(r[18]) or "").upper() == "NEU" else 0,
            "cb": num(r[19]), "pcd": num(r[20]), "h": num(r[21]),
            "et": et, "ets": et_txt,
            "q": num(r[23]) or None, "p": num(r[24]), "kg": num(r[25]),
            "fin": FINITURE.get((s(r[26]) or "").lower()),
        })

    dataset = {
        "meta": {"edizione": edizione, "prezzi": prezzi, "righe": len(righe),
                 "articoli": len({x["c"] for x in righe})},
        "righe": righe,
        "alt": equivalenze(righe),
    }
    return dataset, stats


def portate(righe, codice):
    """Frasi sulla portata scritte nel catalogo per quell'articolo, senza interpretarle."""
    out = []
    for x in righe:
        if x["c"] != codice or not x["n"]:
            continue
        for pezzo in x["n"].split(" - "):
            if RE_PORTATA.search(pezzo) and pezzo not in out:
                out.append(pezzo)
    return out


def leggi_kronprinz():
    from openpyxl import load_workbook
    ws = load_workbook(XLSX_KRONPRINZ, read_only=True, data_only=True).worksheets[0]
    out = {}
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if not s(r[0]):
            continue
        out[s(r[0])] = {"de": s(r[1]), "w": float(r[2]), "d": float(r[3]), "h": int(r[4]),
                        "pcd": float(r[5]), "et": float(r[6]), "cb": float(r[7]),
                        "fin": FINITURE_KRONPRINZ.get(s(r[8]))}
    return out


def norm_pcd(v):
    """Il listino MAK arrotonda gli interassi: 114 sta per 114,3 e 139 per 139,7."""
    return {114.0: 114.3, 139.0: 139.7}.get(float(v), float(v))


def leggi_listino_mak():
    """Codice F... -> misure e finitura dal listino MAK."""
    from openpyxl import load_workbook
    ws = load_workbook(XLSX_LISTINO_MAK, read_only=True, data_only=True).worksheets[0]
    out = {}
    for r in ws.iter_rows(values_only=True, max_col=22):
        if s(r[0]) != "MAK" or not s(r[21]):
            continue
        m = re.match(r"(\d+) X 0*(\d+(?:\.\d+)?)", s(r[15]) or "")
        if m:
            out[s(r[21])] = {"de": None, "w": float(r[12]), "d": float(r[11]), "h": int(m.group(1)),
                             "pcd": float(m.group(2)), "et": float(r[14]), "cb": float(r[16]),
                             "fin": FINITURE_KRONPRINZ.get(s(r[10]))}
    return out


def equivalenze(righe):
    """Mappa R1-... -> alternative approvate. Si ferma se una coppia non torna."""
    kp = leggi_kronprinz()
    listino = leggi_listino_mak()
    with open(CSV_EQUIVALENZE, encoding="utf-8", newline="") as fh:
        coppie = list(csv.DictReader(fh))
    assenti = sorted(set(kp) - {c["codice_alternativo"] for c in coppie})
    if assenti:
        sys.exit(f"Codici Kronprinz senza esito in equivalenze.csv: {assenti}")

    alt = {}
    for c in coppie:
        if c["esito"] != "ok":
            continue
        cod, codice_alt = c["codice_magnetto"], c["codice_alternativo"]
        k = (kp if c["fornitore"] == "Kronprinz" else listino).get(codice_alt)
        art = [x for x in righe if x["c"] == cod]
        if not art or not k:
            sys.exit(f"Coppia con codice inesistente: {cod} / {codice_alt}")
        m = art[0]
        if (float(m["d"]), float(m["wn"]), m["h"], norm_pcd(m["pcd"])) != \
                (k["d"], k["w"], k["h"], norm_pcd(k["pcd"])) or \
                abs(float(m["et"]) - k["et"]) > (0.5 if c["deroga"] == "et" else 0):
            sys.exit(f"Misure diverse: {cod} / {codice_alt}")
        if m["fin"] != k["fin"] and c["deroga"] != "colore":
            sys.exit(f"Finitura diversa senza deroga: {cod} / {codice_alt}")
        if abs(m["cb"] - k["cb"]) > (0.5 if c["deroga"] == "foro" else 0.1) + 1e-6:
            sys.exit(f"Foro centrale diverso: {cod} / {codice_alt}")

        alc_k = set(re.findall(r"\d+", c["alcar"]))
        alc_m = {t for x in art for t in re.findall(r"\d+", x["alc"] or "")}
        w = str(k["w"]).replace(".0", "").replace(".", ",")
        et = str(k["et"]).replace(".0", "").replace(".", ",")
        pcd = str(norm_pcd(k["pcd"])).replace(".0", "").replace(".", ",")
        cb = str(k["cb"]).replace(".0", "").replace(".", ",")
        alt.setdefault(cod, []).append({
            "c": codice_alt, "forn": c["fornitore"],
            "de": k["de"] or (f'ACCIAIO {c["alcar"]}' if c["alcar"] else None),
            "mis": f'{w} x {k["d"]:.0f}" · {k["h"]}x{pcd} · ET {et} · foro {cb}',
            "fin": k["fin"], "por": portate(righe, cod),
            "alcar": None if not alc_k else bool(alc_k & alc_m),
            "nota": c["nota"] or None,
        })
    return alt


def encrypt(plaintext: bytes, pin: str) -> str:
    try:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    except ImportError:
        sys.exit("Manca cryptography: pip3 install cryptography")
    salt = os.urandom(16)
    iv = os.urandom(12)
    key = hashlib.pbkdf2_hmac("sha256", pin.encode("utf-8"), salt, PBKDF2_ITER, dklen=32)
    ct = AESGCM(key).encrypt(iv, plaintext, None)
    b = lambda x: base64.b64encode(x).decode("ascii")
    return b(salt) + "." + b(iv) + "." + b(ct)


def main():
    pin = os.environ.get("MAGNETTO_PIN") or getpass.getpass("PIN Magnetto: ")
    if not pin:
        sys.exit("PIN vuoto: interrompo.")

    dataset, stats = parse()

    with open(PLAIN_JSON, "w", encoding="utf-8") as fh:
        json.dump(dataset, fh, ensure_ascii=False, indent=1)

    plaintext = json.dumps(dataset, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    blob = encrypt(plaintext, pin)
    meta = {"kdf": "PBKDF2-SHA256", "iter": PBKDF2_ITER, "cipher": "AES-GCM"}

    with open(TEMPLATE, encoding="utf-8") as fh:
        tpl = fh.read()
    for token in ("/*DATA_BLOB*/", "/*BUILD_META*/", "/*GENERATO*/"):
        if token not in tpl:
            sys.exit(f"Segnaposto {token} non trovato in index.template.html")
    html = (tpl
            .replace("/*DATA_BLOB*/", json.dumps(blob))
            .replace("/*BUILD_META*/", json.dumps(meta))
            .replace("/*GENERATO*/", json.dumps(dataset["meta"]["edizione"])))
    with open(INDEX, "w", encoding="utf-8") as fh:
        fh.write(html)

    righe = dataset["righe"]
    print("Generati:")
    print("  data/cerchi.plain.json  (NON versionato)")
    print(f"  index.html  ({os.path.getsize(INDEX) / 1024:.0f} KB)")
    print()
    print(f"  edizione catalogo       {dataset['meta']['edizione']}")
    print(f"  prezzi in vigore dal    {dataset['meta']['prezzi']}")
    print(f"  righe valide            {len(righe)}")
    print(f"  articoli unici          {dataset['meta']['articoli']}")
    print(f"  marche                  {len({(x['b'], x['cm']) for x in righe})}")
    print(f"  scartate (vuote)        {stats['vuote']}")
    print(f"  scartate (senza codice) {len(stats['senza_codice'])}  righe {stats['senza_codice']}")
    print(f"  note tradotte dal ted.  {stats['note_tradotte']}")
    print(f"  note rimaste in tedesco {len(stats['note_non_tradotte'])}  "
          f"{sorted(set(stats['note_non_tradotte']))}")
    print(f"  canali non riconosciuti {sorted({x['w'] for x in righe if x['wn'] is None})}")
    print(f"  ET non numerici         {sorted({x['ets'] for x in righe if x['ets']})}")
    print(f"  senza prezzo            {sum(1 for x in righe if x['p'] is None)}")
    alt = dataset["alt"]
    print(f"  Magnetto con alternativa {len(alt)}  "
          f"({len({a['c'] for v in alt.values() for a in v})} codici alternativi)")
    print(f"  Alcar diverso           "
          f"{sorted(c for c, v in alt.items() if any(a['alcar'] is False for a in v))}")


if __name__ == "__main__":
    main()
