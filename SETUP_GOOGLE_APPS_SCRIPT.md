# Configuración del Backend — Google Apps Script

Sigue estos pasos UNA SOLA VEZ para activar la base de datos.

---

## PASO 1 — Crear el Google Sheet

1. Ve a [sheets.google.com](https://sheets.google.com) y crea una hoja nueva.
2. Renombra la hoja principal como `Clientes` (pestaña inferior).
3. En la fila 1, escribe estos encabezados exactamente así (una por columna):

```
A1: id
B1: nombre
C1: email
D1: sesiones_totales
E1: sesiones_usadas
F1: activo
```

4. Crea una segunda pestaña llamada `Reservas` con estos encabezados:

```
A1: id
B1: cliente_id
C1: cliente_nombre
D1: fecha_propuesta
E1: hora_propuesta
F1: mensaje
G1: estado
H1: fecha_creacion
```

5. Crea una tercera pestaña llamada `Config` con:
```
A1: clave
B1: valor
A2: nombre_entrenador
B2: Tu Nombre Aquí
```

6. **Copia el ID de tu Google Sheet** desde la URL:
   `https://docs.google.com/spreadsheets/d/` **ESTE_ES_EL_ID** `/edit`

---

## PASO 2 — Crear el Apps Script

1. En tu Google Sheet ve a **Extensiones → Apps Script**.
2. Borra todo el código que aparece por defecto.
3. Pega este código:

```javascript
const SHEET_ID = 'PEGA_AQUÍ_EL_ID_DE_TU_SHEET';

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const action = e.parameter.action || (e.postData ? JSON.parse(e.postData.contents).action : null);
  const data = e.postData ? JSON.parse(e.postData.contents) : e.parameter;
  
  try {
    let result;
    switch(action) {
      case 'getClientes':       result = getClientes(); break;
      case 'addCliente':        result = addCliente(data); break;
      case 'updateCliente':     result = updateCliente(data); break;
      case 'deleteCliente':     result = deleteCliente(data); break;
      case 'getClienteByEmail': result = getClienteByEmail(data.email); break;
      case 'getReservas':       result = getReservas(); break;
      case 'addReserva':        result = addReserva(data); break;
      case 'updateReserva':     result = updateReserva(data); break;
      case 'usarSesion':        result = usarSesion(data); break;
      default:                  result = { error: 'Acción desconocida' };
    }
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

// ---- CLIENTES ----

function getClientes() {
  const sheet = getSheet('Clientes');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => ({
    id: row[0], nombre: row[1], email: row[2],
    sesiones_totales: row[3], sesiones_usadas: row[4], activo: row[5]
  }));
}

function addCliente(data) {
  const sheet = getSheet('Clientes');
  const id = generateId();
  sheet.appendRow([id, data.nombre, data.email, data.sesiones_totales || 10, 0, true]);
  return { success: true, id };
}

function updateCliente(data) {
  const sheet = getSheet('Clientes');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      if (data.nombre !== undefined)           sheet.getRange(i+1, 2).setValue(data.nombre);
      if (data.email !== undefined)            sheet.getRange(i+1, 3).setValue(data.email);
      if (data.sesiones_totales !== undefined) sheet.getRange(i+1, 4).setValue(data.sesiones_totales);
      if (data.sesiones_usadas !== undefined)  sheet.getRange(i+1, 5).setValue(data.sesiones_usadas);
      if (data.activo !== undefined)           sheet.getRange(i+1, 6).setValue(data.activo);
      return { success: true };
    }
  }
  return { error: 'Cliente no encontrado' };
}

function deleteCliente(data) {
  const sheet = getSheet('Clientes');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Cliente no encontrado' };
}

function getClienteByEmail(email) {
  const clientes = getClientes();
  const cliente = clientes.find(c => c.email.toLowerCase() === email.toLowerCase() && c.activo);
  return cliente || { error: 'Cliente no encontrado' };
}

// ---- RESERVAS ----

function getReservas() {
  const sheet = getSheet('Reservas');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => ({
    id: row[0], cliente_id: row[1], cliente_nombre: row[2],
    fecha_propuesta: row[3], hora_propuesta: row[4],
    mensaje: row[5], estado: row[6], fecha_creacion: row[7]
  }));
}

function addReserva(data) {
  const sheet = getSheet('Reservas');
  const id = generateId();
  sheet.appendRow([
    id, data.cliente_id, data.cliente_nombre,
    data.fecha_propuesta, data.hora_propuesta,
    data.mensaje || '', 'pendiente',
    new Date().toISOString()
  ]);
  return { success: true, id };
}

function updateReserva(data) {
  const sheet = getSheet('Reservas');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      if (data.estado !== undefined) sheet.getRange(i+1, 7).setValue(data.estado);
      return { success: true };
    }
  }
  return { error: 'Reserva no encontrada' };
}

function usarSesion(data) {
  const result = updateCliente({ id: data.cliente_id, sesiones_usadas: data.sesiones_usadas });
  if (data.reserva_id) updateReserva({ id: data.reserva_id, estado: 'completada' });
  return result;
}
```

4. Reemplaza `'PEGA_AQUÍ_EL_ID_DE_TU_SHEET'` con el ID que copiaste en el Paso 1.
5. Guarda el script (Ctrl+S o ⌘+S).

---

## PASO 3 — Publicar como Web App

1. Haz clic en **Implementar → Nueva implementación**.
2. Tipo: **Aplicación web**.
3. Ejecutar como: **Yo (tu cuenta Google)**.
4. Quién tiene acceso: **Cualquier usuario**.
5. Haz clic en **Implementar**.
6. Autoriza los permisos que te pida.
7. **Copia la URL** que aparece — se parece a:
   `https://script.google.com/macros/s/XXXXXXXXXXXX/exec`

---

## PASO 4 — Pegar la URL en los HTML

Abre `admin.html` y `client.html` y busca esta línea al inicio:

```javascript
const API_URL = 'PEGA_AQUÍ_TU_URL_DE_APPS_SCRIPT';
```

Reemplaza el texto por tu URL del Paso 3.

---

## PASO 5 — Subir a GitHub Pages

1. Crea un repositorio en GitHub (puede ser privado o público).
2. Sube los tres archivos: `admin.html`, `client.html`.
3. Ve a **Settings → Pages** y activa GitHub Pages desde la rama `main`.
4. Tu app estará en: `https://TU_USUARIO.github.io/TU_REPO/`
   - Entrenador: `.../admin.html`
   - Cliente: `.../client.html`

---

## Seguridad

- Comparte `admin.html` **solo contigo** (puedes añadir una capa extra con contraseña en el propio HTML si quieres).
- Comparte `client.html` con todos tus clientes.
- El entrenador debe dar de alta el email del cliente antes de que pueda entrar.
