"""
Měsíční výstupy ve stejném formátu, jaký se dnes používá:

1. PND_RRRR_MM.xlsx  – data ČEZ Distribuce (listy „denní“, „hodina“, „1_4_hod“)
2. Vyroba_FVE_RRRR_MM.pdf – report „FVE TCP Šumperk“ (denní bilance hybridní FVE
   z monitoringu + odběr ze sítě podle ČEZ + ocenění výroby k rozúčtování)

Vstupy:
- čtvrthodinová data ČEZ: DataFrame se sloupci
      cas (konec čtvrthodiny), odber [kW], pretok [kW], status
  a volitelně jalova_odber, jalova_dodavka [kVAr]
- export „Energetická bilance – Měsíc“ z monitoringu FVE (xlsx)

Ruční použití:
  python -m app.reporty pnd  <ctvrthodiny.xlsx|csv> 2026 9
  python -m app.reporty pdf  <bilance.xlsx> <odber_cez_kWh> [faktura_sit_Kc]
"""
from __future__ import annotations

import calendar
import html
import os
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Font

DATA = Path(os.environ.get("DATA_DIR", Path(__file__).resolve().parent.parent / "data"))
VYSTUPY = DATA / "vystupy"

STATUS_OK = "naměřená data OK"
DNY = ["pondělí", "úterý", "středa", "čtvrtek", "pátek", "sobota", "neděle"]
MESICE = ["leden", "únor", "březen", "duben", "květen", "červen", "červenec",
          "srpen", "září", "říjen", "listopad", "prosinec"]


# ====================================================================== PND
def cas_pnd(ts: pd.Timestamp) -> str:
    """Konec intervalu ve formátu PND; půlnoc se píše jako 24:00:00 předchozího dne."""
    if ts.hour == 0 and ts.minute == 0:
        return (ts - timedelta(days=1)).strftime("%d.%m.%Y") + " 24:00:00"
    return ts.strftime("%d.%m.%Y %H:%M:%S")


def parsuj_cas_pnd(text) -> pd.Timestamp:
    """Opak cas_pnd – „30.09.2026 24:00:00“ → 2026-10-01 00:00."""
    if isinstance(text, (datetime, pd.Timestamp)):
        return pd.Timestamp(text)
    t = str(text).strip()
    if " 24:00" in t:
        return pd.Timestamp(datetime.strptime(t[:10], "%d.%m.%Y")) + timedelta(days=1)
    for fmt in ("%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M"):
        try:
            return pd.Timestamp(datetime.strptime(t, fmt))
        except ValueError:
            pass
    return pd.Timestamp(t)


def _nejhorsi_status(statusy: pd.Series) -> str:
    """Den/hodina je OK jen když jsou OK všechny čtvrthodiny (stejně jako PND)."""
    jine = [s for s in statusy.dropna() if s != STATUS_OK]
    if not jine:
        return STATUS_OK
    return "neplatná data" if any("neplatn" in s for s in jine) else jine[0]


def pnd_mesic_xlsx(ctvrthodiny: pd.DataFrame, rok: int, mesic: int, cil: Path | None = None) -> Path:
    """Vytvoří PND_RRRR_MM.xlsx se stejnými listy, sloupci a formáty jako export z PND."""
    df = ctvrthodiny.copy()
    df["cas"] = df["cas"].map(parsuj_cas_pnd)
    zacatek = pd.Timestamp(rok, mesic, 1)
    konec = zacatek + pd.offsets.MonthBegin(1)
    # čtvrthodina 00:15 je první v měsíci, 24:00 posledního dne poslední
    df = df[(df["cas"] > zacatek) & (df["cas"] <= konec)].sort_values("cas")
    for c in ("jalova_odber", "jalova_dodavka"):
        if c not in df:
            df[c] = 0.0
    df["interval"] = df["cas"] - pd.Timedelta(minutes=15)  # začátek intervalu
    df["den"] = df["interval"].dt.normalize()
    df["hodina"] = df["interval"].dt.floor("h") + pd.Timedelta(hours=1)  # konec hodiny

    wb = openpyxl.Workbook()
    tucne = Font(bold=True)

    # --- denní
    ws = wb.active
    ws.title = "denní"
    ws.append(["Datum", "Denní činná spotřeba A+ [kWh]", "Denní činná dodávka A- [kWh]",
               "Denní jalový odběr +Ri [kVAr]", "Denní jalová dodávka  -Ri [kVAr]", "Status"])
    dny = df.groupby("den").agg(
        odber=("odber", "sum"), pretok=("pretok", "sum"),
        ri_p=("jalova_odber", "sum"), ri_m=("jalova_dodavka", "sum"),
        status=("status", _nejhorsi_status),
    )
    for den, r in dny.iterrows():
        ws.append([den.to_pydatetime(), r.odber / 4, r.pretok / 4, r.ri_p / 4, r.ri_m / 4, r.status])
    n = ws.max_row
    ws.append([])
    ws.append(["Suma"] + [f"=SUM({c}2:{c}{n})" for c in "BCDE"])
    ws.append(["Účiník", f"=COS(D{n}/B{n})"])  # vzorec převzatý ze vzorového souboru
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[0].number_format = "dd\\.\\ mm/yyyy\\.\\ [$-405]dddd\\ "
        for c in row[1:5]:
            c.number_format = "#,##0.0"
    ws.cell(ws.max_row, 2).number_format = "0.000"
    ws.cell(ws.max_row, 1).number_format = "General"
    ws.cell(ws.max_row - 1, 1).number_format = "General"
    _sirky(ws, {"A": 20.9, "B": 14, "F": 20})

    # --- hodina (průměrný výkon za hodinu)
    ws = wb.create_sheet("hodina")
    ws.append(["Datum", "Hodinová data odběr [kW]", "Hodinová data přetok [kW]",
               "Hodinová data jalová [kVAr]", "Status"])
    hod = df.groupby("hodina").agg(
        odber=("odber", "mean"), pretok=("pretok", "mean"),
        jalova=("jalova_odber", "mean"), status=("status", _nejhorsi_status),
    )
    for h, r in hod.iterrows():
        ws.append([cas_pnd(h), _cislo(r.odber), _cislo(r.pretok), _cislo(r.jalova), r.status])
    _sirky(ws, {"A": 17, "B": 14, "E": 20})

    # --- 1_4_hod
    ws = wb.create_sheet("1_4_hod")
    ws.append(["Datum", "Čtvrthodina odběr [kW]", "Čtvrthodina přetok [kW]", "Status"])
    for r in df.itertuples():
        ws.append([cas_pnd(r.cas), _cislo(r.odber), _cislo(r.pretok), r.status])
    _sirky(ws, {"A": 17, "B": 14, "D": 20})

    for w in wb.worksheets:
        for c in w[1]:
            c.font = tucne

    cil = cil or VYSTUPY / f"PND_{rok}_{mesic:02d}.xlsx"
    cil.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cil)
    return cil


def nacti_pnd_xlsx(cesta: Path) -> pd.DataFrame:
    """Načte list 1_4_hod z exportu PND do DataFrame pro pnd_mesic_xlsx."""
    df = pd.read_excel(cesta, sheet_name="1_4_hod")
    df.columns = ["cas", "odber", "pretok", "status"]
    df["cas"] = df["cas"].map(parsuj_cas_pnd)
    return df


def _cislo(x: float):
    x = round(float(x), 4)
    return int(x) if x == int(x) else x


def _sirky(ws, sirky: dict) -> None:
    for sloupec, s in sirky.items():
        ws.column_dimensions[sloupec].width = s


# ====================================================================== bilance FVE
BILANCE_SLOUPCE = ["obdobi", "prima", "vybijeni", "nabijeni", "odber_site", "dodavka", "spotreba", "vyroba"]


def nacti_bilanci(cesta: Path) -> pd.DataFrame:
    """Export „Energetická bilance – Měsíc“ z monitoringu FVE (hlavička 10 řádků, pak tabulka)."""
    raw = pd.read_excel(cesta, header=None)
    hlavicka = raw.index[raw[0].astype(str).str.strip().eq("Období")][0]
    df = raw.iloc[hlavicka + 1:, :8].dropna(subset=[0]).copy()
    df.columns = BILANCE_SLOUPCE
    df["obdobi"] = pd.to_datetime(df["obdobi"])
    for c in BILANCE_SLOUPCE[1:]:
        df[c] = pd.to_numeric(df[c])
    return df.reset_index(drop=True)


# ====================================================================== PDF
def vyroba_fve_pdf(
    bilance: pd.DataFrame,
    odber_cez_kwh: float,
    cil: Path | None = None,
    cena_fve: float = 7.0,
    koeficient: float = 1.04,
    faktura_sit_kc: float | None = None,
    vypracoval_dne: date | None = None,
    vypracoval: tuple[str, ...] = ("Antonín Londa", "tel. 722 271 261"),
    firma: tuple[str, ...] = ("Agenal s. r. o.", "e mail:", "londa@agenal.cz"),
    nazev: str = "FVE TCP  Šumperk",
) -> Path:
    """Report „FVE TCP Šumperk“ ve stejném rozvržení jako ruční PDF (A4, jedna strana)."""
    prvni = bilance["obdobi"].iloc[0]
    rok, mesic = prvni.year, prvni.month
    suma = bilance[BILANCE_SLOUPCE[1:]].sum()
    k_rozuctovani = suma.vyroba - suma.dodavka  # výroba − dodávka do sítě
    cena_sit = (faktura_sit_kc / odber_cez_kwh) if faktura_sit_kc and odber_cez_kwh else None
    vypracoval_dne = vypracoval_dne or date.today()

    def f1(x):  # 1 desetinné místo, desetinná čárka
        return f"{x:,.1f}".replace(",", " ").replace(".", ",")

    def f0(x):  # zaokrouhlení jako v Excelu (0,5 nahoru), mezera jako oddělovač tisíců
        from decimal import ROUND_HALF_UP, Decimal
        return f"{int(Decimal(str(x)).quantize(Decimal('1'), ROUND_HALF_UP)):,}".replace(",", " ")

    radky = []
    for r in bilance.itertuples():
        d = r.obdobi
        vikend = d.weekday() >= 5
        hodnoty = "".join(f"<td>{f1(getattr(r, c))}</td>" for c in BILANCE_SLOUPCE[1:])
        radky.append(
            f"<tr class='{'vikend' if vikend else ''}'><td class='l'>{d:%d}. {d:%m}.{d:%Y}. {DNY[d.weekday()]}</td>{hodnoty}</tr>"
        )
    sumy = "".join(f"<td>{f0(suma[c])}</td>" for c in BILANCE_SLOUPCE[1:])

    stranka = f"""<!doctype html><html lang="cs"><head><meta charset="utf-8"><style>
@page {{ size: A4; margin: 22mm 18mm 15mm 18mm; }}
body {{ font-family: Calibri, Carlito, Arial, sans-serif; font-size: 10.5pt; color: #000; }}
h1 {{ text-align: center; font-size: 13pt; margin: 0 0 4mm; }}
table {{ border-collapse: collapse; }}
.hl {{ width: 100%; border: 2px solid #000; }}
.hl td, .hl th {{ border: 1px solid #000; padding: 1px 4px; text-align: center; font-weight: normal; }}
.hl td.l {{ text-align: left; }}
.hl tr.vikend td {{ background: #f8cbad; }}
.hl tr.suma td {{ background: #92d050; border-top: 2px solid #000; }}
.box {{ border-collapse: collapse; margin-top: 3mm; table-layout: fixed; }}
.box td.nic {{ background: none; border: 0; }}
.box td {{ border: 1px solid #000; background: #bdd7ee; padding: 1px 4px; white-space: nowrap; }}
.box .cer {{ color: #e00; }} .box .b {{ font-weight: bold; }} .r {{ text-align: right; }}
.pata {{ margin-top: 8mm; display: grid; grid-template-columns: 52mm 33mm auto; row-gap: 1mm; }}
.pata a {{ color: #4f7392; text-decoration: none; }}
</style></head><body>
<h1>{html.escape(nazev)}</h1>
<table class="hl">
<tr><th rowspan="2" style="width:30%">Období</th><th>Přímá<br>spotřeba</th><th>Vybíjení<br>baterie</th><th>Nabíjení<br>baterie</th><th>Odběr ze<br>sítě</th><th>Dodávka<br>do sítě</th><th>Celková<br>spotřeba</th><th>Celková<br>výroba</th></tr>
<tr></tr>
<tr><td>{MESICE[mesic - 1]} {rok}</td>{'<td>[kWh]</td>' * 7}</tr>
{''.join(radky)}
<tr class="suma"><td>Sumary</td>{sumy}</tr>
</table>
<table class="box">
<colgroup><col style="width:46mm"><col style="width:20mm"><col style="width:20mm"><col style="width:20mm"><col style="width:16mm"><col style="width:20mm"><col style="width:12mm"></colgroup>
<tr><td colspan="2">Koeficient * odběr ze sítě</td><td class="r">{f"{koeficient:.2f}".replace(".", ",")}</td><td class="r">{f0(koeficient * suma.odber_site)}</td><td colspan="3" class="nic"></td></tr>
<tr><td colspan="2">Odběr ze sítě měření ČEZ</td><td class="r cer">{f0(odber_cez_kwh)}</td><td class="cer">kWh</td><td colspan="3" class="nic"></td></tr>
<tr><td colspan="4">Fakturace dodavatel síť</td><td></td><td></td><td class="r">{f0(faktura_sit_kc or 0)}&nbsp;Kč</td></tr>
<tr><td colspan="2">Přepočtená cena 1 kWh</td><td class="r">síť</td><td class="r">{(f"{cena_sit:.2f}".replace(".", ",") + "&nbsp;") if cena_sit else ""}Kč</td><td class="r">FVE</td><td class="r cer b">{f"{cena_fve:.2f}".replace(".", ",")}</td><td>Kč</td></tr>
<tr><td colspan="4">Vyrobená energie FVS za měsíc k rozúčtování</td><td class="r cer b">{f0(k_rozuctovani)}</td><td>kWh</td><td class="r"><span class="cer b">{f0(k_rozuctovani * cena_fve)}</span>&nbsp;Kč</td></tr>
</table>
<div class="pata">
<div>Vypracoval&nbsp; {vypracoval_dne:%d.%m.%Y}</div><div>{html.escape(vypracoval[0])}</div><div>{html.escape(firma[0])}</div>
<div></div><div>{html.escape(vypracoval[1])}</div><div>{html.escape(firma[1])} &nbsp; <a>{html.escape(firma[2])}</a></div>
</div>
</body></html>"""

    cil = cil or VYSTUPY / f"Vyroba_FVE_{rok}_{mesic:02d}.pdf"
    cil.parent.mkdir(parents=True, exist_ok=True)
    cil.with_suffix(".html").write_text(stranka, encoding="utf-8")
    _html_do_pdf(stranka, cil)
    return cil


def _html_do_pdf(stranka: str, cil: Path) -> None:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        kwargs = {}
        if os.environ.get("CHROMIUM_PATH"):
            kwargs["executable_path"] = os.environ["CHROMIUM_PATH"]
        browser = p.chromium.launch(**kwargs)
        page = browser.new_page()
        page.set_content(stranka, wait_until="load")
        page.pdf(path=str(cil), format="A4", prefer_css_page_size=True, print_background=True)
        browser.close()


# ====================================================================== měsíc
def mesicni_vystupy(rok: int, mesic: int, bilance: Path | None = None,
                    faktura_sit_kc: float | None = None, koeficient: float = 1.04,
                    cena_fve: float = 7.0) -> dict:
    """PND_RRRR_MM.xlsx z uložených dat ČEZ a (je-li dodána bilance FVE) i PDF."""
    from .stahovac import nacti_ulozena

    data = nacti_ulozena()
    data["cas"] = pd.to_datetime(data["cas"])
    zacatek = pd.Timestamp(rok, mesic, 1)
    konec = zacatek + pd.offsets.MonthBegin(1)
    mesic_df = data[(data["cas"] > zacatek) & (data["cas"] <= konec)]
    if mesic_df.empty:
        raise ValueError(f"Za {mesic:02d}/{rok} nejsou stažená žádná data z ČEZ")
    ocekavano = calendar.monthrange(rok, mesic)[1] * 96
    vysledek = {
        "ctvrthodin": len(mesic_df),
        "ocekavano": ocekavano,
        "odber_cez_kwh": mesic_df["odber"].sum() / 4,
        "pretok_cez_kwh": mesic_df["pretok"].sum() / 4,
        "pnd": pnd_mesic_xlsx(mesic_df, rok, mesic),
    }
    if bilance:
        bil = nacti_bilanci(bilance)
        vysledek["pdf"] = vyroba_fve_pdf(bil, vysledek["odber_cez_kwh"], cena_fve=cena_fve,
                                         koeficient=koeficient, faktura_sit_kc=faktura_sit_kc)
    return vysledek


# ====================================================================== CLI
if __name__ == "__main__":
    prikaz = sys.argv[1] if len(sys.argv) > 1 else ""
    if prikaz == "pnd":
        src = Path(sys.argv[2])
        df = nacti_pnd_xlsx(src) if src.suffix == ".xlsx" else pd.read_csv(src, sep=";", decimal=",")
        print(pnd_mesic_xlsx(df, int(sys.argv[3]), int(sys.argv[4])))
    elif prikaz == "pdf":
        faktura = float(sys.argv[4]) if len(sys.argv) > 4 else None
        print(vyroba_fve_pdf(nacti_bilanci(Path(sys.argv[2])), float(sys.argv[3]), faktura_sit_kc=faktura))
    else:
        print(__doc__)
