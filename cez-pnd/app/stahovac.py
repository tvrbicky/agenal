"""
Stahovač čtvrthodinových dat z Portálu naměřených dat ČEZ Distribuce (PND).

- přihlásí se headless prohlížečem (Playwright) účtem DIP,
- stáhne CSV export za zadané období,
- připojí data do Excelu (list „15min“ + denní souhrn „denne“),
- každý běh a jeho log zapisuje do SQLite (zobrazuje webové rozhraní),
- při chybě uloží screenshot a HTML stránky do data/debug/<id běhu>/.

Selektory portálu jsou v SELEKTORY – při změně portálu se upravují jen tam.

Ruční spuštění:  python -m app.stahovac [YYYY-MM-DD YYYY-MM-DD]
"""
from __future__ import annotations

import os
import sqlite3
import sys
import traceback
from contextlib import closing
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

DATA = Path(os.environ.get("DATA_DIR", Path(__file__).resolve().parent.parent / "data"))
CSV_DIR = DATA / "csv"
DEBUG_DIR = DATA / "debug"
EXCEL = DATA / "cez_data.xlsx"
DB = DATA / "behy.sqlite"

PND_URL = os.environ.get("PND_URL", "https://pnd.cezdistribuce.cz/cezpnd2/external/dashboard/view")

# Kroky portálu ověřené na účtu TCP 1. 11. 2026 (přihlášení přes mepas.cez.cz, dashboard má dvě okna
# .pnd-window; pracujeme v prvním). Postup v okně: Export → Sestava → Množina zařízení → Vlastní období
# → Vyhledat data → Exportovat data → CSV.
SELEKTORY = {
    "cookies": "#CybotCookiebotDialogBodyLevelButtonLevelOptinAllowallSelection",
    "login_user": "input[placeholder='Zadejte svůj e-mail'], input[name='username'], input[type='email']",
    "login_heslo": "input[placeholder='Zadejte své heslo'], input[name='password'], input[type='password']",
    "login_odeslat": "button[type='submit']",
    "po_prihlaseni": ".pnd-window",
    "modal_precteno": ".modal-dialog button:has-text('Přečteno')",
    "okno": ".pnd-window",
    "tlacitko_export": "button[title='Export']",
    "vyhledat": "button:has-text('Vyhledat data')",
    "exportovat": "button:has-text('Exportovat data')",
    "csv": "a:text-is('CSV'):visible",
}
# „1/4 diagram“ = čtvrthodinový odběr i přetok v jednom exportu (stejné sloupce jako list 1_4_hod).
# „Rychlá sestava“ = záložní cesta přes dva profily (+A) a (-A).
SESTAVA = os.environ.get("PND_SESTAVA", "1/4 diagram")
PROFIL_ODBER = os.environ.get("PND_PROFIL_ODBER", "")
PROFIL_DODAVKA = os.environ.get("PND_PROFIL_DODAVKA", "")


# ---------------------------------------------------------------- databáze
def db() -> sqlite3.Connection:
    DATA.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    con.executescript(
        """
        CREATE TABLE IF NOT EXISTS behy (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            zacatek TEXT NOT NULL,
            konec TEXT,
            spustil TEXT,
            obdobi_od TEXT,
            obdobi_do TEXT,
            stav TEXT NOT NULL,          -- bezi | ok | chyba
            radku INTEGER,
            soubor TEXT,
            zprava TEXT
        );
        CREATE TABLE IF NOT EXISTS log (
            beh_id INTEGER NOT NULL,
            cas TEXT NOT NULL,
            uroven TEXT NOT NULL,
            text TEXT NOT NULL
        );
        """
    )
    return con


class Beh:
    """Jeden běh stahování; vše, co se děje, jde do tabulky log."""

    def __init__(self, od: date, do: date, spustil: str):
        self.od, self.do = od, do
        with closing(db()) as con, con:
            cur = con.execute(
                "INSERT INTO behy (zacatek, spustil, obdobi_od, obdobi_do, stav) VALUES (?,?,?,?, 'bezi')",
                (_ted(), spustil, od.isoformat(), do.isoformat()),
            )
            self.id = cur.lastrowid

    def log(self, text: str, uroven: str = "INFO") -> None:
        print(f"[běh {self.id}] {uroven} {text}", flush=True)
        with closing(db()) as con, con:
            con.execute("INSERT INTO log VALUES (?,?,?,?)", (self.id, _ted(), uroven, text))

    def konec(self, stav: str, radku: int | None = None, soubor: str | None = None, zprava: str = "") -> None:
        with closing(db()) as con, con:
            con.execute(
                "UPDATE behy SET konec=?, stav=?, radku=?, soubor=?, zprava=? WHERE id=?",
                (_ted(), stav, radku, soubor, zprava, self.id),
            )


def _ted() -> str:
    return datetime.now().isoformat(timespec="seconds")


# ---------------------------------------------------------------- portál
def _klik(loc, timeout: int = 15_000) -> None:
    """Klik; portál PND rozbíjí Playwrightí kontrolu „hit target“ (kHoverHitTarget…),
    proto při chybě klikneme bez kontrol a nakonec čistým JS."""
    try:
        loc.click(timeout=timeout)
    except Exception:
        try:
            loc.click(force=True, timeout=5_000)
        except Exception:
            loc.evaluate("e => e.click()")


def _pole(okno, popisek: str):
    """Skupina formuláře (label + vue-multiselect) podle přesného textu labelu."""
    return okno.locator(f"label:text-is('{popisek}')").first.locator(
        "xpath=ancestor::div[contains(@class,'form-group')][1]")


def _vybrano(okno, popisek: str) -> str:
    try:
        return (_pole(okno, popisek).locator(".multiselect__single").first.text_content(timeout=2_000) or "").strip()
    except Exception:
        return ""


def _vyber(page, okno, popisek: str, text: str, beh: "Beh", pokusu: int = 3) -> None:
    """Vybere položku ve vue-multiselectu: otevřít, napsat hledaný text, kliknout na shodu (nebo Enter).

    Přednost má přesná shoda („ELM 90121971“ před „Elektroměr VT TCP 90121971“).
    """
    if text in _vybrano(okno, popisek):
        beh.log(f"{popisek}: „{_vybrano(okno, popisek)}“ (už vybráno)")
        return
    pole = _pole(okno, popisek)
    for _ in range(20):  # portál pole odemyká až po načtení sestavy
        if not pole.locator(".multiselect--disabled").count():
            break
        page.wait_for_timeout(500)
    else:
        beh.log(f"Pole „{popisek}“ je v portálu zamčené, pokračuji s „{_vybrano(okno, popisek)}“ "
                f"(chtěli jsme „{text}“) – výsledek ověřte", "WARN")
        return
    for pokus in range(pokusu):
        _klik(pole.locator(".multiselect").first)
        page.wait_for_timeout(400)
        vstup = pole.locator("input.multiselect__input").first
        if vstup.count():
            try:
                vstup.fill(text, timeout=5_000)
            except Exception:
                vstup.press_sequentially(text, delay=30)
            page.wait_for_timeout(600)
        moznosti = _nabidka(page, popisek)
        if moznosti and not any(text in m for m in moznosti):
            raise RuntimeError(f"V poli „{popisek}“ není „{text}“; nabídka: {moznosti}")
        presne = [m for m in moznosti if m in (text, f"ELM {text}")]
        cil = presne[0] if presne else next((m for m in moznosti if text in m), text)
        volba = pole.locator(".multiselect__option:visible").filter(has_text=cil).first
        if volba.count():
            _klik(volba)
        elif vstup.count():
            vstup.press("Enter")
        page.wait_for_timeout(600)
        if text in _vybrano(okno, popisek):
            beh.log(f"{popisek}: „{_vybrano(okno, popisek)}“")
            return
        beh.log(f"{popisek}: pokus {pokus + 1} nevyšel (vybráno „{_vybrano(okno, popisek)}“)", "WARN")
        page.keyboard.press("Escape")
    raise RuntimeError(f"Nepodařilo se vybrat „{text}“ v poli „{popisek}“")


def _texty(page, selektor: str) -> list[str]:
    """Texty prvků čistým JS – Playwrightí all_inner_texts na portálu PND padá
    (TypeError: kHoverHitTargetInterceptorEvents is not iterable)."""
    try:
        return [t.strip() for t in page.evaluate(
            "s => Array.prototype.map.call(document.querySelectorAll(s), e => e.textContent || '')", selektor)
            if t and t.strip()]
    except Exception:
        return []


def _nabidka(page, popisek: str) -> list[str]:
    """Všechny položky multiselectu v prvním okně (čistým JS, bez otevírání nabídky)."""
    try:
        return [t.strip() for t in page.evaluate(
            """p => { const w = document.querySelector('.pnd-window'); if (!w) return [];
                const l = Array.prototype.find.call(w.querySelectorAll('label'), x => x.textContent.trim() === p);
                const g = l && l.closest('.form-group'); if (!g) return [];
                return Array.prototype.map.call(g.querySelectorAll('.multiselect__option'), e => e.textContent || ''); }""",
            popisek) if t and t.strip()]
    except Exception:
        return []


def _najdi_profil(odkazy: list[str], zadany: str, znak: str) -> str:
    if zadany:
        return zadany
    for t in odkazy:
        if f"({znak})" in t and "za den" not in t.lower():
            return t
    raise RuntimeError(f"Mezi profily není čtvrthodinový ({znak}); dostupné: {odkazy}. Nastavte PND_PROFIL_*.")


def _export_csv(page, okno, cil: Path, beh: "Beh") -> Path:
    _klik(okno.locator(SELEKTORY["exportovat"]).first)
    page.wait_for_timeout(500)
    formaty = _texty(page, ".pnd-window .dropdown-menu a")
    if formaty:
        beh.log("Formáty exportu: " + ", ".join(formaty))
    with page.expect_download(timeout=180_000) as dl:
        _klik(page.locator(SELEKTORY["csv"]).first)
    dl.value.save_as(cil)
    beh.log(f"Staženo {cil.name} ({cil.stat().st_size} B)")
    return cil


def stahni_csv(beh: Beh) -> list[Path]:
    """Stáhne export za období beh.od–beh.do. Vrací 1 soubor (1/4 diagram) nebo 2 (profily +A, -A)."""
    from playwright.sync_api import sync_playwright

    user = os.environ["CEZ_USER"]
    heslo = os.environ["CEZ_PASSWORD"]
    elm = os.environ.get("CEZ_ELM", "").strip()
    CSV_DIR.mkdir(parents=True, exist_ok=True)
    debug = DEBUG_DIR / str(beh.id)
    obdobi = f"{beh.od:%d.%m.%Y} 00:00 - {beh.do + timedelta(days=1):%d.%m.%Y} 00:00"
    zaklad = f"{beh.od:%Y%m%d}_{beh.do:%Y%m%d}_beh{beh.id}"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(accept_downloads=True, locale="cs-CZ", viewport={"width": 1600, "height": 1000})
        page.set_default_timeout(30_000)
        try:
            beh.log(f"Otevírám {PND_URL}")
            page.goto(PND_URL, wait_until="networkidle")
            if page.locator(SELEKTORY["cookies"]).count():
                _klik(page.locator(SELEKTORY["cookies"]))
            _snimek(page, debug, "01_prihlaseni")
            beh.log(f"Přihlašovací stránka: {page.url.split('?')[0]}")

            beh.log(f"Přihlašuji se jako {user}")
            page.fill(SELEKTORY["login_user"], user)
            page.fill(SELEKTORY["login_heslo"], heslo)
            _klik(page.locator(SELEKTORY["login_odeslat"]).first)
            page.wait_for_selector(SELEKTORY["po_prihlaseni"], timeout=60_000)
            if page.locator(SELEKTORY["modal_precteno"]).count():
                beh.log("Zavírám oznámení portálu (Přečteno)")
                _klik(page.locator(SELEKTORY["modal_precteno"]).first)
                page.reload(wait_until="networkidle")
                page.wait_for_selector(SELEKTORY["po_prihlaseni"], timeout=60_000)
            page.wait_for_load_state("networkidle")
            _snimek(page, debug, "02_po_prihlaseni")
            beh.log(f"Přihlášení OK ({page.url})")

            okno = page.locator(SELEKTORY["okno"]).first
            _klik(okno.locator(SELEKTORY["tlacitko_export"]).first)
            page.wait_for_timeout(1_000)
            _vyber(page, okno, "Sestava", SESTAVA, beh)
            nabidka = _nabidka(page, "Množina zařízení")
            beh.log("Množiny zařízení / elektroměry v účtu: " + (" | ".join(nabidka) or "(nepodařilo se načíst)"))
            if elm:
                _vyber(page, okno, "Množina zařízení", elm, beh)
            else:
                beh.log(f"Množina zařízení: „{_vybrano(okno, 'Množina zařízení')}“ (CEZ_ELM nevyplněno)", "WARN")
            try:
                _vyber(page, okno, "Období", "Vlastní", beh, pokusu=2)
            except Exception as e:
                beh.log(f"Období „Vlastní“ nevybráno ({e}), vyplňuji rovnou pole Vlastní období", "WARN")
            pole = _pole(okno, "Vlastní období").locator("input").first
            if not pole.count():
                pole = okno.locator("label:text-is('Vlastní období')").locator("xpath=following::input[1]")
            pole.fill(obdobi)
            pole.press("Enter")
            page.wait_for_timeout(500)
            beh.log(f"Vlastní období: {pole.input_value()}")
            _snimek(page, debug, "03_filtr")
            _klik(okno.locator(SELEKTORY["vyhledat"]).first)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(2_000)
            _snimek(page, debug, "04_vysledek")

            if "rychlá" not in SESTAVA.lower():
                return [_export_csv(page, okno, CSV_DIR / f"pnd_{zaklad}.csv", beh)]

            odkazy = [t for t in _texty(page, ".pnd-window a") if "Profil" in t or "(+A)" in t or "(-A)" in t]
            beh.log("Dostupné profily: " + " | ".join(odkazy))
            soubory = []
            for znak, zadany, nazev in (("+A", PROFIL_ODBER, "odber"), ("-A", PROFIL_DODAVKA, "dodavka")):
                profil = _najdi_profil(odkazy, zadany, znak)
                beh.log(f"Stahuji profil „{profil}“")
                _klik(okno.locator("a", has_text=profil).first)
                page.wait_for_load_state("networkidle")
                _snimek(page, debug, f"05_{nazev}")
                soubory.append(_export_csv(page, okno, CSV_DIR / f"pnd_{nazev}_{zaklad}.csv", beh))
            return soubory
        except Exception:
            _snimek(page, debug, "chyba")
            beh.log(f"Screenshot a HTML stránky uloženy do debug/{beh.id}", "WARN")
            raise
        finally:
            browser.close()


def _snimek(page, slozka: Path, nazev: str) -> None:
    """Uloží screenshot + HTML – slouží k ladění selektorů."""
    if os.environ.get("DEBUG_SNIMKY", "1") != "1" and nazev != "chyba":
        return
    try:
        slozka.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=slozka / f"{nazev}.png", full_page=True)
        (slozka / f"{nazev}.html").write_text(page.content(), encoding="utf-8")
    except Exception:
        pass


# ---------------------------------------------------------------- Excel
SLOUPCE = ["cas", "odber", "pretok", "status"]


def _precti(cesta: Path) -> pd.DataFrame:
    """Načte export z PND (CSV se středníky v UTF-8/cp1250, nebo XLSX s listem 1_4_hod)."""
    if cesta.suffix.lower() in (".xlsx", ".xls"):
        xls = pd.ExcelFile(cesta)
        df = pd.read_excel(xls, sheet_name="1_4_hod" if "1_4_hod" in xls.sheet_names else 0)
    else:
        df = None
        for kodovani in ("utf-8-sig", "cp1250"):
            try:
                df = pd.read_csv(cesta, sep=";", decimal=",", encoding=kodovani)
                break
            except UnicodeDecodeError:
                continue
    if df is None or df.empty:
        raise ValueError(f"Soubor {cesta.name} je prázdný nebo nečitelný")
    return df


def _pojmenuj(df: pd.DataFrame, vychozi: str | None = None) -> pd.DataFrame:
    """Sloupce podle názvu (+A/odběr, -A/přetok/dodávka, status); první sloupec je čas.

    `vychozi` = jak pojmenovat jediný hodnotový sloupec exportu jednoho profilu,
    pokud se z názvu nepozná (např. „Hodnota [kW]“).
    """
    mapa = {df.columns[0]: "cas"}
    for c in df.columns[1:]:
        n = str(c).lower()
        if "status" in n:
            mapa[c] = "status"
        elif "+a" in n or "odběr" in n or "spotřeb" in n:
            mapa[c] = "odber"
        elif "-a" in n or "přetok" in n or "dodávk" in n or "výrob" in n:
            mapa[c] = "pretok"
    if vychozi and vychozi not in mapa.values():
        zbyle = [c for c in df.columns[1:] if c not in mapa]
        if zbyle:
            mapa[zbyle[0]] = vychozi
    return df.rename(columns=mapa)


def _cas_a_cisla(df: pd.DataFrame, sloupce: list[str]) -> pd.DataFrame:
    from .reporty import parsuj_cas_pnd

    df = df.copy()
    df["cas"] = df["cas"].map(lambda t: parsuj_cas_pnd(t) if pd.notna(t) and str(t).strip() else pd.NaT)
    df = df.dropna(subset=["cas"])
    for c in sloupce:
        if df[c].dtype == object:
            df[c] = df[c].astype(str).str.replace("\xa0", "").str.replace(" ", "").str.replace(",", ".")
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def nacti_csv(cesta: Path) -> pd.DataFrame:
    """Jeden export s oběma směry (+A i -A) → cas, odber, pretok, status.

    Čas „24:00:00“ se převede na půlnoc dalšího dne.
    """
    df = _pojmenuj(_precti(cesta))
    chybi = [c for c in ("odber", "pretok") if c not in df]
    if chybi:
        raise ValueError(f"V exportu chybí sloupce {chybi}, nalezeno: {list(df.columns)}")
    if "status" not in df:
        df["status"] = "naměřená data OK"
    return _cas_a_cisla(df, ["odber", "pretok"])[SLOUPCE]


def nacti_profily(odber: Path, dodavka: Path) -> pd.DataFrame:
    """Dva samostatné exporty profilů (+A a -A), jak je dává PND → jedna tabulka."""
    a = _pojmenuj(_precti(odber), "odber")
    b = _pojmenuj(_precti(dodavka), "pretok")
    for df, sl, f in ((a, "odber", odber), (b, "pretok", dodavka)):
        if sl not in df:
            raise ValueError(f"V {f.name} chybí hodnoty ({sl}), nalezeno: {list(df.columns)}")
    a = _cas_a_cisla(a, ["odber"])
    b = _cas_a_cisla(b, ["pretok"])
    st_a = a["status"] if "status" in a else "naměřená data OK"
    st_b = b.set_index("cas")["status"] if "status" in b else None
    vysl = a.assign(status=st_a)[["cas", "odber", "status"]].merge(
        b[["cas", "pretok"]], on="cas", how="outer")
    if st_b is not None:
        # horší z obou statusů (OK jen když jsou OK oba)
        sb = vysl["cas"].map(st_b)
        vysl["status"] = [x if (y is None or pd.isna(y) or y == "naměřená data OK") else y
                          for x, y in zip(vysl["status"].fillna("chybí data +A"), sb)]
    vysl["status"] = vysl["status"].fillna("chybí data +A")
    return vysl.sort_values("cas")[SLOUPCE]


def nacti_ulozena() -> pd.DataFrame:
    """Všechna dosud stažená čtvrthodinová data (list 15min v cez_data.xlsx)."""
    if not EXCEL.exists():
        return pd.DataFrame(columns=SLOUPCE)
    df = pd.read_excel(EXCEL, sheet_name="15min")
    return df.rename(columns={"konec čtvrthodiny": "cas", "odběr A+ [kW]": "odber",
                              "přetok A- [kW]": "pretok"})[SLOUPCE]


def uloz_do_excelu(nova: pd.DataFrame) -> int:
    """Připojí data do cez_data.xlsx bez duplicit; vrací počet řádků v souboru."""
    data = pd.concat([nacti_ulozena(), nova]).drop_duplicates(subset="cas", keep="last").sort_values("cas")

    # 15min výkon v kW -> energie za den v kWh = součet / 4 (den podle začátku intervalu)
    den = (pd.to_datetime(data["cas"]) - pd.Timedelta(minutes=15)).dt.normalize()
    denni = data.groupby(den)[["odber", "pretok"]].sum() / 4
    denni.columns = ["odběr A+ [kWh]", "přetok A- [kWh]"]

    EXCEL.parent.mkdir(parents=True, exist_ok=True)
    tmp = EXCEL.with_suffix(".tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        data.rename(columns={"cas": "konec čtvrthodiny", "odber": "odběr A+ [kW]",
                             "pretok": "přetok A- [kW]"}).to_excel(xw, sheet_name="15min", index=False)
        denni.rename_axis("den").reset_index().to_excel(xw, sheet_name="denne", index=False)
    tmp.replace(EXCEL)
    return len(data)


# ---------------------------------------------------------------- běh
def spust(od: date | None = None, do: date | None = None, spustil: str = "plán", csv: Path | None = None) -> int:
    """Celý běh. `csv` = zpracovat existující soubor místo stahování (test)."""
    od = od or date.today() - timedelta(days=1)
    do = do or od
    beh = Beh(od, do, spustil)
    try:
        if csv:
            soubory = [csv]
            df = nacti_csv(csv)
        else:
            soubory = stahni_csv(beh)
            df = nacti_csv(soubory[0]) if len(soubory) == 1 else nacti_profily(*soubory)
        beh.log(f"Načteno {len(df)} čtvrthodin ({df['cas'].min()} – {df['cas'].max()}), "
                f"odběr {df['odber'].sum() / 4:.2f} kWh, dodávka {df['pretok'].sum() / 4:.2f} kWh")
        nok = df[df["status"] != "naměřená data OK"]
        if len(nok):
            beh.log(f"{len(nok)} čtvrthodin nemá status OK (např. {nok['status'].iloc[0]})", "WARN")
        celkem = uloz_do_excelu(df)
        beh.log(f"Excel aktualizován, celkem {celkem} řádků")
        beh.konec("ok", radku=len(df), soubor=", ".join(f.name for f in soubory))
    except Exception as e:
        beh.log(traceback.format_exc(), "ERROR")
        beh.konec("chyba", zprava=f"{type(e).__name__}: {e}")
    return beh.id


if __name__ == "__main__":
    args = [date.fromisoformat(a) for a in sys.argv[1:3]]
    spust(*args, spustil="ručně (CLI)")
