# Setup del formulario de la landing LíderMente

## Paso 1 — Crear la hoja de Google Sheets

1. Ve a [sheets.new](https://sheets.new)
2. Nómbrala **"LíderMente — Leads Landing"**
3. En la fila 1 (encabezados), escribe exactamente:

| A | B | C | D | E |
|---|---|---|---|---|
| Fecha | Nombre | WhatsApp | Negocio | Origen |

4. Copia el **ID** de la hoja (el código largo en la URL entre `/d/` y `/edit`).

---

## Paso 2 — Crear el Apps Script

1. En la hoja, ve a **Extensiones → Apps Script**
2. Borra todo el código y pega esto:

```javascript
function doPost(e) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var datos = JSON.parse(e.postData.contents);

  var fila = hoja.getLastRow() + 1;
  hoja.getRange(fila, 1).setValue(datos.fecha || new Date().toISOString());
  hoja.getRange(fila, 2).setValue(datos.nombre || "");
  hoja.getRange(fila, 3).setValue(datos.whatsapp || "");
  hoja.getRange(fila, 3).setNumberFormat('@');  // Proteger el número como texto
  hoja.getRange(fila, 4).setValue(datos.negocio || "");
  hoja.getRange(fila, 5).setValue(datos.origen || "landing");

  return ContentService.createTextOutput(
    JSON.stringify({ status: "ok", fila: fila })
  ).setMimeType(ContentService.MimeType.JSON);
}
```

3. Guarda (Ctrl+S), ponle nombre **"WebApp Leads"**

---

## Paso 3 — Publicar como Web App

1. Click en **Implementar → Nueva implementación**
2. Tipo: **Aplicación web**
3. Descripción: "Leads LíderMente landing"
4. Ejecutar como: **Yo (tu correo)**
5. Quién tiene acceso: **Cualquier persona**
6. Click en **Implementar**
7. Copia la URL que te da (empieza con `https://script.google.com/macros/s/...`)

---

## Paso 4 — Conectar a la landing

Abre `index.html` y busca esta línea:

```
const ENDPOINT = "https://script.google.com/macros/s/PLACEHOLDER/exec";
```

Reemplaza `PLACEHOLDER` con el ID real de tu implementación.

---

## Verificación

1. Abre la landing en el navegador
2. Llena el formulario con datos de prueba
3. Verifica que aparezcan en la hoja de Google Sheets
4. Si no llegan, ve a Apps Script → Ejecuciones para ver errores

---

## Deploy en Netlify

1. Ve a [app.netlify.com/drop](https://app.netlify.com/drop)
2. Arrastra la carpeta `lidermente/` completa (index.html + que-se-vive.mp4)
3. Configura el dominio personalizado si aplica
