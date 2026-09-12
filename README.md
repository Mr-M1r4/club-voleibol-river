# 🏐 Club Voleibol River — Portal web

Portal oficial del Club Deportivo de Voleibol River (Rionegro, Antioquia).
Este sitio es **estático y está siempre disponible** (GitHub Pages). Conecta con el
sistema de gestión que vive en **Google Apps Script**.

Estructura:

- `index.html` — portal de acceso (asistencia QR, dashboard, QR de afiliado, encuesta)
- `encuesta_voley.html` — encuesta de estudio de mercado (guarda en Google Sheets)
- `config.js` — **configuración central**: pega aquí tu URL del Web App

---

## Paso 1 · Desplegar el sistema en Google Apps Script

El motor (afiliados, pagos, asistencia por QR, formularios y notificaciones) usa APIs de
Google y debe ejecutarse en Apps Script. El proyecto completo está en la carpeta `src/`
(o en `River.json`) del mismo directorio de descarga.

### Opción A — con clasp (recomendado)

```bash
npm install -g @google/clasp
clasp login

# dentro de la carpeta con el código (src/)
clasp create --type webapp --title "Club Voleibol River"
clasp push
clasp open
```

### Opción B — manual

1. Crea un proyecto en https://script.google.com (o abre el Google Sheet del club → Extensiones → Apps Script).
2. Crea cada archivo de `src/` con su mismo nombre: `Code.gs`, `Base.gs`, `Afiliados.gs`,
   `Pagos.gs`, `Asistencia.gs`, `QR.gs`, `Notificaciones.gs`, `Forms.gs`, y las páginas
   `Index.html`, `Encuesta.html` (más `appsscript.json` en Configuración del proyecto).
3. Pega el contenido de cada archivo.

### Publicar como Web App

1. En el editor: **Implementar → Nueva implementación → Aplicación web**.
2. Configura:
   - **Ejecutar como:** *Yo* (`Mr-M1r4` / tu cuenta)
   - **Quién tiene acceso:** *Cualquier persona*
3. Copia la URL que termina en `/exec`.
4. Ejecuta una vez **`inicializarSistema()`** desde el menú `🏐 Club Voleibol River`
   (o en el editor) para crear las pestañas del Sheet.
5. Desde el menú ejecuta **`Generar QR para todos los afiliados`** para que los QR
   apunten a la URL definitiva.

> Cada vez que edites el código y quieras publicar cambios usa
> **Implementar → Administrar implementaciones → Editar → Nueva versión**.
> La URL `/exec` publicada no cambia.

---

## Paso 2 · Conectar el portal

Edita `config.js`:

```js
window.CLUB_VOLEIBOL = {
  appUrl: "https://script.google.com/macros/s/AKfycb.../exec"
};
```

Listo: el portal, la encuesta y el buscador de QR quedarán conectados al sistema.

---

## Paso 3 · (Opcional) Proteger el webhook de pagos

En la pestaña **Config** del Google Sheet pon un valor en `WEBHOOK_TOKEN`. Desde ese
momento los POST a `?action=webhook_pago` deberán incluir
`?token=TU_TOKEN` (o header `x-api-key`) para ser aceptados.

---

## Notas de seguridad

- Mantén privada la pestaña **Config** del Sheet (contiene tokens de Telegram/Twilio).
- La URL `/exec` compartida da acceso a *asistencia/dashboard QR/encuesta*; el resto de
  acciones administrativas siguen controladas por editores del Sheet.