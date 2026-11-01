"""
Webové rozhraní: historie běhů, logy, screenshoty z ladění, ruční spuštění,
stažení Excelu. Zároveň obsahuje plánovač denního stahování.

Spuštění:  python -m app.web   (port 8080, nebo proměnná PORT)
"""
from __future__ import annotations

import os
import threading
from contextlib import closing
from datetime import date, datetime, timedelta
from pathlib import Path
from werkzeug.utils import secure_filename
from functools import wraps

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from flask import Flask, Response, abort, flash, redirect, render_template, request, send_from_directory, url_for

from . import reporty
from . import stahovac as s

app = Flask(__name__)
app.jinja_env.filters["todt"] = datetime.fromisoformat
app.secret_key = os.environ.get("WEB_PASSWORD") or os.urandom(16)
ZAMEK = threading.Lock()
PLAN = os.environ.get("PLAN_CRON", "30 6 * * *")  # min hod den měsíc den_v_týdnu
TZ = os.environ.get("TZ", "Europe/Prague")


# ---------------------------------------------------------------- přihlášení
def chraneno(f):
    @wraps(f)
    def obal(*a, **kw):
        user, heslo = os.environ.get("WEB_USER"), os.environ.get("WEB_PASSWORD")
        if user and heslo:
            auth = request.authorization
            if not auth or auth.username != user or auth.password != heslo:
                return Response("Přihlaste se", 401, {"WWW-Authenticate": 'Basic realm="CEZ PND"'})
        return f(*a, **kw)
    return obal


# ---------------------------------------------------------------- běhy
def spust_na_pozadi(od=None, do=None, spustil="plán") -> bool:
    if not ZAMEK.acquire(blocking=False):
        return False

    def prace():
        try:
            s.spust(od, do, spustil=spustil)
        finally:
            ZAMEK.release()

    threading.Thread(target=prace, daemon=True).start()
    return True


planovac = BackgroundScheduler(timezone=TZ)
planovac.add_job(spust_na_pozadi, CronTrigger.from_crontab(PLAN, timezone=TZ), id="denni")


def pnd_za_minuly_mesic():
    """2. den v měsíci vytvoří PND_RRRR_MM.xlsx za předchozí měsíc."""
    minuly = date.today().replace(day=1) - timedelta(days=1)
    try:
        reporty.mesicni_vystupy(minuly.year, minuly.month)
    except Exception as e:
        print(f"Měsíční PND se nepodařilo vytvořit: {e}", flush=True)


planovac.add_job(pnd_za_minuly_mesic, CronTrigger(day=2, hour=8, timezone=TZ), id="mesicni")


# ---------------------------------------------------------------- stránky
@app.get("/")
@chraneno
def prehled():
    with closing(s.db()) as con:
        behy = con.execute("SELECT * FROM behy ORDER BY id DESC LIMIT 100").fetchall()
    job = planovac.get_job("denni")
    excel = s.EXCEL.stat() if s.EXCEL.exists() else None
    return render_template(
        "prehled.html",
        behy=behy,
        dalsi=getattr(job, "next_run_time", None),
        plan=PLAN,
        excel=excel,
        bezi=ZAMEK.locked(),
        vcera=date.fromordinal(date.today().toordinal() - 1).isoformat(),
        chybi=[k for k in ("CEZ_USER", "CEZ_PASSWORD") if not os.environ.get(k)],
        vystupy=sorted((p for p in reporty.VYSTUPY.glob("*") if p.suffix in (".xlsx", ".pdf")),
                       key=lambda p: p.stat().st_mtime, reverse=True) if reporty.VYSTUPY.exists() else [],
        minuly=(date.today().replace(day=1) - timedelta(days=1)).strftime("%Y-%m"),
    )


@app.get("/beh/<int:beh_id>")
@chraneno
def detail(beh_id: int):
    with closing(s.db()) as con:
        beh = con.execute("SELECT * FROM behy WHERE id=?", (beh_id,)).fetchone()
        log = con.execute("SELECT * FROM log WHERE beh_id=? ORDER BY rowid", (beh_id,)).fetchall()
    if not beh:
        abort(404)
    slozka = s.DEBUG_DIR / str(beh_id)
    snimky = sorted(p.name for p in slozka.glob("*")) if slozka.exists() else []
    return render_template("detail.html", beh=beh, log=log, snimky=snimky)


@app.post("/spustit")
@chraneno
def spustit():
    od = date.fromisoformat(request.form["od"]) if request.form.get("od") else None
    do = date.fromisoformat(request.form["do"]) if request.form.get("do") else od
    spust_na_pozadi(od, do, spustil="ručně (web)")
    return redirect(url_for("prehled"))


@app.post("/mesicni")
@chraneno
def mesicni():
    rok, mesic = map(int, request.form["mesic"].split("-"))
    bilance = None
    soubor = request.files.get("bilance")
    if soubor and soubor.filename:
        nahrano = s.DATA / "nahrano"
        nahrano.mkdir(parents=True, exist_ok=True)
        bilance = nahrano / f"bilance_{rok}_{mesic:02d}.xlsx"
        soubor.save(bilance)
    faktura = request.form.get("faktura", "").replace(" ", "").replace(",", ".")
    koef = request.form.get("koeficient", "1.04").replace(",", ".")
    try:
        v = reporty.mesicni_vystupy(rok, mesic, bilance, float(faktura) if faktura else None, float(koef))
        cz = lambda x: f"{x:,.1f}".replace(",", " ").replace(".", ",")
        zprava = (f"{mesic:02d}/{rok}: odběr ČEZ {cz(v['odber_cez_kwh'])} kWh, přetok {cz(v['pretok_cez_kwh'])} kWh, "
                  f"{v['ctvrthodin']} z {v['ocekavano']} čtvrthodin. Vytvořeno: {v['pnd'].name}"
                  + (f", {v['pdf'].name}" if "pdf" in v else " (PDF se vytvoří po nahrání bilance FVE)"))
        flash(zprava, "ok")
    except Exception as e:
        flash(f"Chyba: {e}", "chyba")
    return redirect(url_for("prehled"))


@app.post("/nahrat-pnd")
@chraneno
def nahrat_pnd():
    """Ruční nahrání exportu z PND (CSV/XLSX) – záloha, když automatické stažení selže."""
    soubor = request.files["pnd"]
    s.CSV_DIR.mkdir(parents=True, exist_ok=True)
    cil = s.CSV_DIR / ("rucne_" + secure_filename(soubor.filename))
    soubor.save(cil)
    beh_id = s.spust(date.today(), spustil="ručně (nahrán soubor)", csv=cil)
    return redirect(url_for("detail", beh_id=beh_id))


@app.get("/vystup/<path:nazev>")
@chraneno
def vystup(nazev: str):
    return send_from_directory(reporty.VYSTUPY, nazev, as_attachment=True)


@app.get("/excel")
@chraneno
def excel():
    return send_from_directory(s.DATA, s.EXCEL.name, as_attachment=True)


@app.get("/csv/<path:nazev>")
@chraneno
def csv(nazev: str):
    return send_from_directory(s.CSV_DIR, nazev, as_attachment=True)


@app.get("/debug/<int:beh_id>/<path:nazev>")
@chraneno
def debug(beh_id: int, nazev: str):
    return send_from_directory(s.DEBUG_DIR / str(beh_id), nazev)


@app.get("/zdravi")
def zdravi():
    return {"ok": True, "bezi": ZAMEK.locked()}


def main():
    from waitress import serve

    s.db().close()
    planovac.start()
    port = int(os.environ.get("PORT", 8080))
    print(f"Web na portu {port}, plán stahování: {PLAN} ({TZ})", flush=True)
    serve(app, host="0.0.0.0", port=port)


if __name__ == "__main__":
    main()
