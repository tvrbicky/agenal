# Stahování dat z ČEZ Distribuce (PND) – VPS

Každý den stáhne čtvrthodinová data o odběru (+A) a dodávce do sítě (-A)
z Portálu naměřených dat ČEZ Distribuce a připojí je do `data/cez_data.xlsx`
(list `15min` a denní souhrn v kWh `denne`). Webové rozhraní ukazuje historii běhů,
log každého běhu, screenshoty z portálu a umožní ruční stažení libovolného období.

## Instalace na VPS (Docker)

```bash
# 1. nahrát složku cez-pnd na VPS, např. do /opt/cez-pnd
cd /opt/cez-pnd
cp .env.example .env && nano .env      # vyplnit CEZ_USER, CEZ_PASSWORD, CEZ_ELM, WEB_PASSWORD
chmod 600 .env

# 2. spustit
docker compose up -d --build
docker compose logs -f
```

Web běží na `127.0.0.1:8080`. Ven ho vystavte jen přes reverzní proxy s HTTPS,
např. Caddy (`/etc/caddy/Caddyfile`):

```
cez.vasedomena.cz {
    reverse_proxy 127.0.0.1:8080
}
```

## První spuštění a ladění

Kroky v portálu (`app/stahovac.py`, slovník `SELEKTORY`) jsou převzaté z ověřené komunitní
integrace ondrejvysek/HomeAssistant-CEZDistribuce-PND: přihlášení → okno Export → Sestava
„Rychlá sestava“ → Množina zařízení = `CEZ_ELM` → Období „Vlastní“ → Vyhledat data → Tabulka dat
→ profil (+A) a profil (−A) → Exportovat data → CSV. Stahují se dva CSV, ty se spojí podle času.
Na našem účtu je potřeba je jednou ověřit – log běhu vypíše dostupné elektroměry i názvy profilů;
kdyby automatický výběr čtvrthodinového profilu nesedl, nastavte `PND_PROFIL_ODBER`/`PND_PROFIL_DODAVKA`.
Import do Excelu je otestovaný na zářijových datech (výstup se vzorem PND shoduje na 100 %).

1. Ve webu klikněte na **Stáhnout** (včerejšek).
2. Otevřete detail běhu: log + screenshot a HTML každého kroku (`DEBUG_SNIMKY=1`).
3. Podle snímků upravte `SELEKTORY`/kroky, `docker compose up -d --build`, opakujte.
4. Až běh projde, nastavte `DEBUG_SNIMKY=0` (snímky se pak ukládají jen při chybě).

Historii (až 36 měsíců zpět) stáhnete ručně zadáním období ve webu.

## Měsíční výstupy (Excel PND + PDF „FVE TCP Šumperk“)

Ve webu v kartě **Měsíční výstupy** vyberte měsíc, nahrajte export
„Energetická bilance – Měsíc“ z monitoringu FVE a volitelně částku faktury za síť.
Vzniknou dva soubory (stáhnete je v kartě Vytvořené soubory):

- `PND_RRRR_MM.xlsx` – listy `denní`, `hodina`, `1_4_hod` ve stejném formátu jako export z PND,
- `Vyroba_FVE_RRRR_MM.pdf` – stejné rozvržení jako dosavadní ruční PDF; řádek
  „Odběr ze sítě měření ČEZ“ se plní ze stažených dat ČEZ (A+), přepočtená cena sítě
  = faktura / odběr ČEZ, výroba k rozúčtování = celková výroba − dodávka do sítě.

`PND_RRRR_MM.xlsx` za minulý měsíc se navíc vytváří automaticky každého 2. dne v měsíci.
Kdyby automatické stažení z portálu selhalo, lze export z PND nahrát ručně (karta Nahrát export z PND).

Pozn.: sloupce jalové energie (+Ri/−Ri) zatím zůstávají nulové, stahuje se jen činná energie A+ a A−.

## Struktura

| Cesta | Obsah |
|---|---|
| `app/stahovac.py` | přihlášení do PND, export CSV, zápis do Excelu, log do SQLite |
| `app/reporty.py` | měsíční Excel PND a PDF report výroby FVE |
| `data/vystupy/` | vytvořené měsíční Excely a PDF |
| `app/web.py` | webové rozhraní + plánovač (`PLAN_CRON`, výchozí 6:30 denně) |
| `data/cez_data.xlsx` | výsledný Excel |
| `data/csv/` | originální CSV z každého běhu |
| `data/behy.sqlite` | historie běhů a logy |
| `data/debug/<běh>/` | screenshoty a HTML z portálu |

Ruční běh bez webu: `docker compose exec cez-pnd python -m app.stahovac 2026-10-01 2026-10-31`
