# Setup del cuestionario de Onboarding (`/onboarding/`)

Cada envío hace tres cosas:

1. Guarda una fila en Google Sheets con todas las respuestas.
2. Crea una carpeta por cliente en Google Drive, dentro de **2. AGENCIA / Onboarding clientes** (`Empresa — fecha`), con los archivos adjuntos y un `.txt` con las respuestas.
3. Te manda un correo con el resumen y el enlace a la carpeta.

Tiempo estimado: 10 minutos.

---

## Paso 1 — Abrir la hoja (ya está creada)

La carpeta y la hoja ya existen en tu Drive:

- Carpeta: **2. AGENCIA / Onboarding clientes** → https://drive.google.com/drive/folders/18wTMb5dDbNt2qwlq8yYpTBF09WPE0bTK
- Hoja: **Onboarding clientes — Respuestas** → https://docs.google.com/spreadsheets/d/1jwr8Myl55gqvRoaAZaCVYcYHQQ2sqXef2AZuAu00lYM/edit

Déjala vacía: los encabezados se crean solos con el primer envío.

---

## Paso 2 — Pegar el Apps Script

1. En la hoja, ve a **Extensiones → Apps Script**
2. Borra todo y pega esto:

```javascript
const CARPETA_ID = "18wTMb5dDbNt2qwlq8yYpTBF09WPE0bTK"; // 2. AGENCIA / Onboarding clientes
const AVISAR_A = "hola@leilafittipaldi.com";      // Correo que recibe el aviso de cada envío
const ZONA = "America/Cancun";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    const r = d.respuestas || {};
    const empresa = r.empresa || d.cliente || "Sin nombre";
    const fecha = Utilities.formatDate(new Date(), ZONA, "yyyy-MM-dd HH:mm");

    // 1) Carpeta con adjuntos
    const carpeta = DriveApp.getFolderById(CARPETA_ID).createFolder(empresa + " — " + fecha);
    (d.archivos || []).forEach(function (a) {
      carpeta.createFile(Utilities.newBlob(Utilities.base64Decode(a.base64), a.tipo, a.nombre));
    });
    carpeta.createFile("Respuestas - " + empresa + ".txt", d.texto || "", MimeType.PLAIN_TEXT);

    // 2) Fila en la hoja (los encabezados se agregan solos si hay preguntas nuevas)
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const fijos = ["Fecha", "Empresa", "Nombre", "Correo", "WhatsApp", "Carpeta", "Enlaces", "Archivos"];
    let enc = hoja.getLastColumn() ? hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0] : [];
    if (!enc.length) enc = fijos.slice();
    Object.keys(r).forEach(function (k) { if (enc.indexOf(k) === -1) enc.push(k); });
    hoja.getRange(1, 1, 1, enc.length).setValues([enc]).setFontWeight("bold");
    hoja.setFrozenRows(1);

    const base = {
      Fecha: fecha, Empresa: empresa, Nombre: r.nombre, Correo: r.email, WhatsApp: r.whatsapp,
      Carpeta: carpeta.getUrl(),
      Enlaces: (d.enlaces || []).join("\n"),
      Archivos: (d.archivos || []).map(function (a) { return a.nombre; }).join("\n")
    };
    const fila = enc.map(function (h) { return String((h in base ? base[h] : r[h]) || ""); });
    const n = hoja.getLastRow() + 1;
    hoja.getRange(n, 1, 1, fila.length).setNumberFormat("@").setValues([fila]);

    // 3) Aviso por correo
    MailApp.sendEmail({
      to: AVISAR_A,
      replyTo: r.email || AVISAR_A,
      subject: "Nuevo onboarding: " + empresa,
      body: "Carpeta con adjuntos: " + carpeta.getUrl() + "\n\n" + (d.texto || "")
    });

    return json_({ status: "ok" });
  } catch (err) {
    return json_({ status: "error", message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return json_({ status: "ok", servicio: "onboarding" }); }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
```

3. Guarda (Ctrl+S) con el nombre **"WebApp Onboarding"**.

---

## Paso 3 — Publicar como Web App

1. **Implementar → Nueva implementación**
2. Tipo: **Aplicación web**
3. Ejecutar como: **Yo (tu correo)**
4. Quién tiene acceso: **Cualquier persona**
5. **Implementar** → autoriza los permisos (Drive, Sheets y Gmail; aparece "Google no verificó esta app" → *Configuración avanzada → Ir a WebApp Onboarding*).
6. Copia la URL (`https://script.google.com/macros/s/.../exec`).

> Si después cambias el código: **Implementar → Administrar implementaciones → editar → Nueva versión**. Así la URL no cambia.

---

## Paso 4 — Conectar la página

En `onboarding/index.html` busca:

```
endpoint: "https://script.google.com/macros/s/PLACEHOLDER/exec",
```

y pega tu URL. Haz commit y push.

---

## Verificación

1. Abre `https://leilafittipaldi.com/onboarding/?cliente=Prueba`
2. Llénalo con datos de prueba y adjunta un PDF pequeño.
3. Revisa: fila nueva en la hoja, subcarpeta en **2. AGENCIA / Onboarding clientes** con el PDF y correo de aviso.
4. Si no llega nada: Apps Script → **Ejecuciones** para ver el error.

## Enlace para cada cliente

`https://leilafittipaldi.com/onboarding/?cliente=Nombre%20del%20cliente`

Ejemplo: `https://leilafittipaldi.com/onboarding/?cliente=Capitalízate%20by%20Gretel%20García`

Límites: 10 MB por archivo y 20 MB en total por envío (cambiar en `CONFIG` si hace falta; Apps Script acepta hasta ~50 MB por petición).
