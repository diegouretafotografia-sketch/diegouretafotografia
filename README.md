# Diego Ureta Fotografía — sitio web

- `public/index.html` — la web completa (una sola página).
- `wrangler.jsonc` — configuración para publicar en Cloudflare (Workers con assets estáticos).
- `apps-script/Code.gs` — receptor del formulario de cotización (Google Apps Script → Google Sheet + correo).

Cada cambio subido a la rama `main` se publica automáticamente en Cloudflare.
