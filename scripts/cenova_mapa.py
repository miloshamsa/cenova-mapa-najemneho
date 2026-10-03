#!/usr/bin/env python3
"""Stáhne nejnovější cenovou mapu nájemného MF ČR a převede ji na JSON pro kalkulačku.

Zdroj: https://mf.gov.cz/cs/rozpoctova-politika/podpora-projektoveho-rizeni/cenova-mapa/cenova-mapa-infografika
MF vydává XLSX čtvrtletně (do 45 dnů po konci čtvrtletí). Skript najde nejnovější soubor,
ověří strukturu a zapíše:

  data/index.json      - seznam lokalit pro vyhledávání + metadata
  data/kraj-<NN>.json  - nájmy po katastrech daného kraje

Když se struktura souboru změní, skript skončí chybou (exit 1) a nic nepřepíše -
GitHub Action pak spadne a pošle e-mail, místo aby web tiše ukazoval nesmysly.

Použití:
  python scripts/cenova_mapa.py                # stáhne nejnovější soubor
  python scripts/cenova_mapa.py --soubor x.xlsx  # zpracuje lokální soubor
"""
from __future__ import annotations

import argparse
import io
import json
import re
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
import requests

ZDROJ_STRANKA = "https://mf.gov.cz/cs/rozpoctova-politika/podpora-projektoveho-rizeni/cenova-mapa/cenova-mapa-infografika"
ZAKLAD_URL = "https://mf.gov.cz"
UA = {"User-Agent": "Mozilla/5.0 (compatible; investrio-cenova-mapa/1.0)"}
VYSTUP = Path(__file__).resolve().parent.parent / "data"

OCEKAVANE_SLOUPCE = [
    "Kraj", "Katastrální území", "Obec", "Kód obce", "VK",
    "Nájemné referenčního bytu za m² v Kč za 1 měsíc",
    "Dolní interval nájemného u referenčního bytu za m² v Kč za 1 měsíc",
    "Horní interval nájemného u referenčního bytu za m² v Kč za 1 měsíc",
    "Nájemné referenčního bytu novostavby za m² v Kč za 1 měsíc",
    "Minimální hodnota nájemného za m² v Kč",
    "Maximální hodnota nájemného za m² v Kč",
    "Mediánová hodnota nájemného za m² v Kč",
    "Datová pokrytost",
]
SIRKA_BLOKU = 10  # 9 hodnot + prázdný oddělovač
MIN_RADKU = 7000  # srpen 2026: 7 631 územních jednotek
ROZSAH_KC_M2 = (40, 2500)

# Pořadí krajů = kód v JSON (kraj-01.json ...); nový kraj v datech = chyba, ne tichý nový kód.
KRAJE = [
    "Hlavní město Praha", "Středočeský kraj", "Jihočeský kraj", "Plzeňský kraj",
    "Karlovarský kraj", "Ústecký kraj", "Liberecký kraj", "Královéhradecký kraj",
    "Pardubický kraj", "Kraj Vysočina", "Jihomoravský kraj", "Olomoucký kraj",
    "Zlínský kraj", "Moravskoslezský kraj",
]
ATRIBUTY = ["balkon", "terasa", "vybaveni", "garaz", "vytah", "jiny_material"]


class ChybaStruktury(Exception):
    pass


def log(msg: str) -> None:
    print(msg, flush=True)


def najdi_nejnovejsi() -> tuple[str, str]:
    """Vrátí (url, název souboru) nejnovější cenové mapy."""
    r = requests.get(ZDROJ_STRANKA, headers=UA, timeout=60)
    r.raise_for_status()
    odkazy = set(re.findall(r'/assets/attachments/(\d{4}-\d{2}-\d{2})_Cenova-mapa(?:_v(\d+))?\.xlsx', r.text))
    if not odkazy:
        raise ChybaStruktury("Na stránce MF nebyl nalezen žádný soubor *_Cenova-mapa*.xlsx")
    datum, verze = max(odkazy, key=lambda x: (x[0], int(x[1] or 0)))
    nazev = f"{datum}_Cenova-mapa{'_v' + verze if verze else ''}.xlsx"
    return f"{ZAKLAD_URL}/assets/attachments/{nazev}", nazev


def cislo(v) -> int | None:
    if v is None or v == "":
        return None
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        raise ChybaStruktury(f"Nečíselná hodnota v datech: {v!r}")


def text_radku(r) -> str:
    return " ".join(str(x) for x in r if x is not None)


def parsuj_info(ws) -> dict:
    """Z listu 'Základní informace' vytáhne období, výměry a příplatky za atributy."""
    radky = [tuple(r) for r in ws.iter_rows(values_only=True)]
    info: dict = {}

    for r in radky:
        t = text_radku(r)
        m = re.search(r"z období od (\d{1,2}\.\d{1,2}\.\d{4}) do (\d{1,2}\.\d{1,2}\.\d{4})", t)
        if m:
            info["obdobi"] = {"od": m.group(1), "do": m.group(2)}
        m = re.search(r"ID:\s*([\d.]+)", t)
        if m:
            info["id"] = m.group(1)

    def tabulka(nadpis_re: str, pocet: int) -> list[list]:
        for i, r in enumerate(radky):
            if re.search(nadpis_re, text_radku(r)):
                hodnoty = [[x for x in rr if x is not None] for rr in radky[i + 2:i + 2 + pocet]]
                return hodnoty
        raise ChybaStruktury(f"Na listu Základní informace chybí tabulka: {nadpis_re}")

    vymery = {}
    for row in tabulka(r"Medián výměr nájemných bytů", len(KRAJE)):
        if len(row) != 5 or row[0] not in KRAJE:
            raise ChybaStruktury(f"Neočekávaný řádek v tabulce výměr: {row}")
        vymery[f"{KRAJE.index(row[0]) + 1:02d}"] = [cislo(x) for x in row[1:]]
    info["vymery_m2"] = vymery

    def priplatky(nadpis_re: str, sloupcu: int) -> dict:
        out = {}
        for row in tabulka(nadpis_re, 4):
            if len(row) != 1 + sloupcu or not str(row[0]).startswith("VK"):
                raise ChybaStruktury(f"Neočekávaný řádek v tabulce příplatků: {row}")
            out[str(row[0])[-1]] = dict(zip(ATRIBUTY, [cislo(x) for x in row[1:]]))
        return out

    info["priplatky"] = {
        "starsi": priplatky(r"pro referenční byt\s+v Kč/m²", 5),
        "novostavba": priplatky(r"pro referenční byt novostavba", 6),
    }
    return info


def parsuj_mapu(ws) -> list[dict]:
    it = ws.iter_rows(values_only=True)
    hlavicka = list(next(it))
    if [h for h in hlavicka[:len(OCEKAVANE_SLOUPCE)]] != OCEKAVANE_SLOUPCE:
        raise ChybaStruktury(f"Změnila se hlavička listu s mapou: {hlavicka[:len(OCEKAVANE_SLOUPCE)]}")
    lokality = []
    for r in it:
        if not r or r[2] is None:
            continue
        kraj = r[0]
        if kraj not in KRAJE:
            raise ChybaStruktury(f"Neznámý kraj: {kraj!r}")
        vk: dict[str, list] = {}
        for b in range(4):
            o = 4 + b * SIRKA_BLOKU
            if o >= len(r) or r[o] is None:
                continue
            k = cislo(r[o])
            ref, lo, hi, nov, mn, mx, med, pokr = (cislo(x) for x in r[o + 1:o + 9])
            for nazev, v in (("ref", ref), ("novostavba", nov)):
                if v is None or not ROZSAH_KC_M2[0] <= v <= ROZSAH_KC_M2[1]:
                    raise ChybaStruktury(f"{r[2]}/{r[1]} VK{k}: {nazev} = {v} mimo rozsah")
            # [ref, dolní, horní, novostavba, min, max, medián, pokrytí]
            vk[str(k)] = [ref, lo, hi, nov, mn, mx, med, pokr]
        if not vk:
            continue
        lokality.append({
            "kraj": f"{KRAJE.index(kraj) + 1:02d}",
            # u obcí bez ORP je jednotkou celá obec a katastr MF nevyplňuje
            "ku": r[1] or r[2],
            "obec": r[2],
            "kod": cislo(r[3]),
            "vk": vk,
        })
    if len(lokality) < MIN_RADKU:
        raise ChybaStruktury(f"Jen {len(lokality)} lokalit (čekám aspoň {MIN_RADKU})")
    return lokality


def zapis(lokality: list[dict], info: dict, nazev: str, url: str | None) -> None:
    meta = {
        "zdroj": "Ministerstvo financí ČR - cenová mapa nájemného",
        "zdroj_url": ZDROJ_STRANKA,
        "soubor": nazev,
        "soubor_url": url,
        "vydano": nazev[:10],
        "zpracovano": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "kraje": {f"{i + 1:02d}": k for i, k in enumerate(KRAJE)},
        "pole_vk": ["ref", "dolni", "horni", "novostavba", "min", "max", "median", "pokryti"],
        **info,
    }
    tmp = Path(tempfile.mkdtemp())
    # index: [ku, obec, kraj] - pořadí = id lokality v souboru kraje
    index, po_krajich = [], {}
    for lok in lokality:
        seznam = po_krajich.setdefault(lok["kraj"], [])
        index.append([lok["ku"], lok["obec"], lok["kraj"], len(seznam)])
        seznam.append([lok["ku"], lok["obec"], lok["kod"], lok["vk"]])
    (tmp / "index.json").write_text(
        json.dumps({"meta": meta, "lokality": index}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    for kraj, seznam in po_krajich.items():
        (tmp / f"kraj-{kraj}.json").write_text(
            json.dumps(seznam, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    # přepis až po úspěšném vygenerování všeho
    VYSTUP.mkdir(parents=True, exist_ok=True)
    for f in VYSTUP.glob("*.json"):
        f.unlink()
    for f in tmp.iterdir():
        f.replace(VYSTUP / f.name)
    log(f"[OK] Zapsáno {len(lokality)} lokalit do {VYSTUP} ({len(po_krajich)} krajů)")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--soubor", help="lokální XLSX místo stažení")
    ap.add_argument("--vynutit", action="store_true", help="přegenerovat i beze změny souboru MF")
    args = ap.parse_args()
    try:
        if args.soubor:
            nazev, url = Path(args.soubor).name, None
            data = Path(args.soubor).read_bytes()
        else:
            url, nazev = najdi_nejnovejsi()
            index_soubor = VYSTUP / "index.json"
            if index_soubor.exists() and not args.vynutit:
                stary = json.loads(index_soubor.read_text(encoding="utf-8"))["meta"].get("soubor")
                if stary == nazev:
                    log(f"[OK] Beze změny, data už jsou ze souboru {nazev}")
                    return 0
            log(f"[..] Stahuji {url}")
            r = requests.get(url, headers=UA, timeout=120)
            r.raise_for_status()
            data = r.content
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        for list_ in ("Základní informace", "Cenové mapy nájemného"):
            if list_ not in wb.sheetnames:
                raise ChybaStruktury(f"Chybí list {list_!r}, listy: {wb.sheetnames}")
        info = parsuj_info(wb["Základní informace"])
        lokality = parsuj_mapu(wb["Cenové mapy nájemného"])
        zapis(lokality, info, nazev, url)
        return 0
    except ChybaStruktury as e:
        log(f"[CHYBA] Struktura souboru MF se změnila: {e}")
        return 1
    except requests.RequestException as e:
        log(f"[CHYBA] Stažení se nepovedlo: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
