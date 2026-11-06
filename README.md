# grupoduralco-web

Sitio web de Grupo Duralco: https://www.grupoduralco.com

- `public/`: todos los archivos del sitio (HTML, scripts, imágenes en `public/assets/`).
- `wrangler.jsonc`: configuración del Cloudflare Worker `muddy-sea-6c08`.

Cada cambio que llega a la rama `main` se publica automáticamente desde Cloudflare (Workers Builds).
