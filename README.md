# AGENAL – energie TCP Šumperk

- [`cez-pnd/`](cez-pnd/) – automatické stahování čtvrthodinových dat z Portálu naměřených dat ČEZ Distribuce,
  webové rozhraní s logem stahování, měsíční Excel PND a PDF „Výroba FVE“. Běží na https://energie.golferis.cz.

## Nasazení na VPS (ai.golferis.cz)

Předpoklad: Docker s `docker compose` a reverzní proxy (Caddy nebo nginx). DNS `energie.golferis.cz` míří na VPS.

```bash
sudo git clone https://github.com/tvrbicky/agenal.git /opt/agenal   # privátní repo: přihlášení tokenem nebo deploy klíčem
cd /opt/agenal/cez-pnd
sudo cp .env.example .env && sudo nano .env    # CEZ_USER, CEZ_PASSWORD, CEZ_ELM, WEB_PASSWORD
sudo chmod 600 .env
sudo docker compose up -d --build
sudo docker compose logs -f                    # Ctrl+C ukončí jen výpis
```

Proxy s HTTPS:
- **Caddy:** obsah [`cez-pnd/deploy/Caddyfile`](cez-pnd/deploy/Caddyfile) přidat do `/etc/caddy/Caddyfile`, `sudo systemctl reload caddy` (certifikát si Caddy vyřídí sám).
- **nginx:** [`cez-pnd/deploy/nginx-energie.conf`](cez-pnd/deploy/nginx-energie.conf) + `certbot --nginx -d energie.golferis.cz`.

Pokud port 8080 na VPS už něco používá, změňte v `cez-pnd/docker-compose.yml` levou stranu `127.0.0.1:8080:8080` (např. `127.0.0.1:8090:8080`) a stejný port v proxy.

První test: na https://energie.golferis.cz klikněte **Stáhnout** (včerejšek) a otevřete detail běhu – log ukáže
elektroměry v účtu a názvy profilů, screenshoty jednotlivých kroků jsou v detailu. Když vše projde,
nastavte v `.env` `DEBUG_SNIMKY=0` a `sudo docker compose up -d`.

Aktualizace: `cd /opt/agenal && sudo git pull && cd cez-pnd && sudo docker compose up -d --build`
