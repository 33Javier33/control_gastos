/**
 * GOOGLE APPS SCRIPT - CONTROL FINANCIERO MULTI-USUARIO
 * Repositorio: 33Javier33/control_gastos
 *
 * INSTRUCCIONES DE DESPLIEGUE:
 * 1. Abre script.google.com y pega este código completo
 * 2. Implementar > Nueva implementación > Aplicación web
 * 3. Ejecutar como: Yo
 * 4. Quién tiene acceso: Cualquier persona
 * 5. Copia la URL generada y pégala en index.html (constante WEB_APP_URL)
 */

// ── INICIALIZACIÓN DE HOJAS ──────────────────────────────────────────────────

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabs = [
    { name: 'Usuarios',    headers: ['RUT', 'Nombre', 'Password'] },
    { name: 'Movimientos', headers: ['Usuario', 'ID', 'Fecha', 'Concepto', 'Monto', 'Categoría', 'Origen_Ingreso', 'Es_Ahorro'] },
    { name: 'Deudas',      headers: ['Usuario', 'ID', 'Fecha', 'Concepto', 'Monto'] },
    { name: 'Categorias',  headers: ['Usuario', 'Tipo', 'Nombre', 'Orden'] },
    { name: 'Archivos',    headers: ['Usuario', 'ID', 'Periodo', 'JSON_Data'] }
  ];
  tabs.forEach(tab => {
    let sheet = ss.getSheetByName(tab.name);
    if (!sheet) {
      sheet = ss.insertSheet(tab.name);
      sheet.getRange(1, 1, 1, tab.headers.length)
           .setValues([tab.headers])
           .setFontWeight('bold')
           .setBackground('#f3f3f3');
      sheet.setFrozenRows(1);
    }
  });
}

// ── GET: LOGIN Y CARGA DE DATOS ──────────────────────────────────────────────

function doGet(e) {
  setupSheet();
  const action = e.parameter.action || '';
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const json   = (obj) => ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);

  // ── Autenticación ──
  if (action === 'login') {
    const rut  = (e.parameter.rut  || '').toUpperCase();
    const pass = (e.parameter.pass || '');
    const users = ss.getSheetByName('Usuarios').getDataRange().getValues().slice(1);
    const user  = users.find(r => String(r[0]).toUpperCase() === rut && String(r[2]) === pass);
    return user ? json({ ok: true, nombre: user[1] }) : json({ ok: false });
  }

  // ── Verificar RUT para recuperación de PIN ──
  if (action === 'checkRut') {
    const rut   = (e.parameter.rut || '').toUpperCase();
    const users = ss.getSheetByName('Usuarios').getDataRange().getValues().slice(1);
    const user  = users.find(r => String(r[0]).toUpperCase() === rut);
    return user ? json({ ok: true, nombre: user[1] }) : json({ ok: false });
  }

  // ── Restablecer PIN ──
  if (action === 'resetPass') {
    const rut     = (e.parameter.rut     || '').toUpperCase();
    const newPass = (e.parameter.newPass || '');
    if (!/^\d{4}$/.test(newPass)) return json({ ok: false, error: 'PIN debe ser 4 dígitos' });
    const sh     = ss.getSheetByName('Usuarios');
    const data   = sh.getDataRange().getValues();
    const rowIdx = data.slice(1).findIndex(r => String(r[0]).toUpperCase() === rut);
    if (rowIdx === -1) return json({ ok: false, error: 'RUT no registrado' });
    sh.getRange(rowIdx + 2, 3).setValue(newPass);
    return json({ ok: true });
  }

  // ── Datos del usuario ──
  if (action === 'data') {
    const rut    = (e.parameter.rut || '').toUpperCase();
    const byUser = (name) => ss.getSheetByName(name)
      .getDataRange().getValues().slice(1)
      .filter(r => String(r[0]).toUpperCase() === rut);

    const cats = byUser('Categorias');
    return json({
      txs: byUser('Movimientos').map(r => ({
        id: r[1], date: r[2], text: r[3], amount: r[4],
        cat: r[5], source: r[6], isSaving: r[7] === 'SI'
      })),
      debts: byUser('Deudas').map(r => ({
        id: r[1], date: r[2], text: r[3], amount: r[4]
      })),
      archives: byUser('Archivos').map(r => ({
        id: r[1], period: r[2], data: JSON.parse(r[3])
      })),
      catsInc: cats.filter(r => r[1] === 'INC').sort((a,b) => a[3]-b[3]).map(r => r[2]),
      catsExp: cats.filter(r => r[1] === 'EXP').sort((a,b) => a[3]-b[3]).map(r => r[2]),
      catsSav: cats.filter(r => r[1] === 'SAV').sort((a,b) => a[3]-b[3]).map(r => r[2])
    });
  }

  return json({});
}

// ── POST: REGISTRO Y GUARDADO ────────────────────────────────────────────────

function doPost(e) {
  setupSheet();
  const body   = JSON.parse(e.postData.contents);
  const action = body.action || '';
  const ss     = SpreadsheetApp.getActiveSpreadsheet();

  // ── Registro de nuevo usuario ──
  if (action === 'register') {
    const rut    = (body.rut || '').toUpperCase();
    const shUsers = ss.getSheetByName('Usuarios');
    const exists  = shUsers.getDataRange().getValues().slice(1)
      .some(r => String(r[0]).toUpperCase() === rut);
    if (!/^\d{4}$/.test(body.pass)) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'PIN debe ser 4 dígitos' })).setMimeType(ContentService.MimeType.JSON);
    if (!exists) shUsers.appendRow([rut, body.nombre, body.pass]);
    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
  }

  // ── Guardar datos del usuario ──
  if (action === 'save') {
    const rut  = (body.rut || '').toUpperCase();
    const data = body.data;

    // Reescribe las filas del usuario en cada hoja conservando las de otros usuarios
    const rewrite = (sheetName, headers, newRows) => {
      const sh      = ss.getSheetByName(sheetName);
      const others  = sh.getDataRange().getValues().slice(1)
        .filter(r => String(r[0]).toUpperCase() !== rut);
      sh.clearContents();
      sh.getRange(1, 1, 1, headers.length)
        .setValues([headers]).setFontWeight('bold').setBackground('#f3f3f3');
      const combined = others.concat(newRows);
      if (combined.length) sh.getRange(2, 1, combined.length, headers.length).setValues(combined);
    };

    rewrite(
      'Movimientos',
      ['Usuario','ID','Fecha','Concepto','Monto','Categoría','Origen_Ingreso','Es_Ahorro'],
      data.txs.map(t => [rut, t.id, t.date, t.text, t.amount, t.cat||'', t.source||'', t.isSaving?'SI':'NO'])
    );

    rewrite(
      'Deudas',
      ['Usuario','ID','Fecha','Concepto','Monto'],
      data.debts.map(d => [rut, d.id, d.date, d.text, d.amount])
    );

    const catRows = [];
    (data.catsInc || []).forEach((c,i) => catRows.push([rut,'INC',c,i]));
    (data.catsExp || []).forEach((c,i) => catRows.push([rut,'EXP',c,i]));
    (data.catsSav || []).forEach((c,i) => catRows.push([rut,'SAV',c,i]));
    rewrite('Categorias', ['Usuario','Tipo','Nombre','Orden'], catRows);

    rewrite(
      'Archivos',
      ['Usuario','ID','Periodo','JSON_Data'],
      (data.archives || []).map(a => [rut, a.id, a.period, JSON.stringify(a.data)])
    );
  }

  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}
