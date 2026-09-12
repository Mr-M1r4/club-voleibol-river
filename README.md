# 🐆 Blue Leopards Volleyball Club — Portal web

Sitio oficial del **Blue Leopards Volleyball Club** (Rionegro, Antioquia), construido con la
identidad del brief oficial de marca (v1.0 · Septiembre 2026).

Este sitio es **estático y está siempre disponible** (GitHub Pages). Actualmente muestra
**información del club** y la **encuesta de estudio de mercado**.

## Contenido

| Archivo            | Descripción                                             |
|--------------------|---------------------------------------------------------|
| `index.html`       | Página principal: el club, identidad, paleta y uniformes |
| `encuesta_voley.html` | Encuesta de clases de vóleibol (guarda en Google Sheets) |
| `config.js`        | Configuración central (`appUrl` del Web App)            |
| `img/`             | Logotipos (Blue Leopards + opciones)                    |

## Conectar la encuesta (dato obligatorio)

La encuesta guarda las respuestas en Google Sheets a través del **Web App de Apps Script**.
Para habilitarla:

1. Sube el código de la carpeta `src/` (o importa `River.json`) en
   https://script.google.com y **publica como aplicación web**:
   - Ejecutar como: *Yo*
   - Quién tiene acceso: *Cualquier persona*
2. Copia la URL que termina en `/exec`.
3. Pégalo en `config.js`:

```js
window.CLUB_VOLEIBOL = {
  appUrl: "https://script.google.com/macros/s/AKfycb.../exec"
};
```

> Sin `appUrl`, el resto del sitio funciona igual: la encuesta se envía pero llega un aviso
> en la consola del navegador y la respuesta no se guarda.

### Despliegue del proyecto con clasp

```bash
npm install -g @google/clasp
clasp login
cd src/            # carpeta con el código (Code.gs, Base.gs, ... Index.html, Encuesta.html)
clasp create --type webapp --title "Blue Leopards Volleyball Club"
clasp push
clasp open
```

Después de publicar, ejecuta **`inicializarSistema()`** y **`Generar QR para todos los afiliados`**
desde el menú `🏐 Club Voleibol River`.

## Notas de identidad

- Paleta oficial: Azul Noche `#12161A` · Azul Leopardo `#70D6FF` · Plata Glaciar `#A9C2D2` · Blanco Nieve `#F4F8FA`.
- Logotipo principal: `img/logo.jpg`. Las opciones alternativas (`logo_1`, `logo_2`, `logo_3`)
  corresponden a las versiones op2/op3/op4; se cambia en `index.html` si se prefiere otra.