/**
 * SCRIPT DE MIGRACIÓN — Google Sheets → Supabase
 *
 * INSTRUCCIONES:
 * 1. Abre script.google.com → abre el proyecto de tu planilla de Control Financiero
 * 2. Crea un archivo nuevo (Archivo > Nuevo > Script), nómbralo MigrateToSupabase
 * 3. Pega este código completo
 * 4. Selecciona "migrateToSupabase" en el menú y haz clic en ▶ Ejecutar
 * 5. Acepta los permisos si te los pide
 * 6. Revisa Ver > Registros de ejecución para ver el resultado
 *
 * El script puede ejecutarse múltiples veces sin problema.
 * Si falla, los datos en Sheets NO se tocan — solo Supabase.
 */

function migrateToSupabase() {
  const API_URL          = 'https://lpulmjzboogixbdxxayo.supabase.co/functions/v1/api';
  const MIGRATION_SECRET = 'mig_Xq7Kp3Nz2Wv9Rm4Jb6Ys';

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const usuarios    = ss.getSheetByName('Usuarios').getDataRange().getValues().slice(1);
  const movimientos = ss.getSheetByName('Movimientos').getDataRange().getValues().slice(1);
  const deudas      = ss.getSheetByName('Deudas').getDataRange().getValues().slice(1);
  const categorias  = ss.getSheetByName('Categorias').getDataRange().getValues().slice(1);
  const archivos    = ss.getSheetByName('Archivos').getDataRange().getValues().slice(1);

  let ok = 0, errors = 0;

  for (const user of usuarios) {
    if (!user[0]) continue;

    const rut           = String(user[0]).toUpperCase();
    const nombre        = String(user[1]);
    const password_hash = String(user[2]);

    // Movimientos
    const txsRaw = movimientos.filter(r => String(r[0]).toUpperCase() === rut);
    const txs = txsRaw.map((r, i) => ({
      id:       String(r[1]) || ('gs_' + rut + '_tx_' + i),  // genera ID si está vacío
      date:     r[2] ? String(r[2]) : '',
      text:     String(r[3] || ''),
      amount:   Number(r[4]) || 0,
      cat:      String(r[5] || ''),
      source:   String(r[6] || ''),
      isSaving: String(r[7]) === 'SI'
    }));

    // Deudas
    const userDebts = deudas
      .filter(r => String(r[0]).toUpperCase() === rut)
      .map((r, i) => ({
        id:     String(r[1]) || ('gs_' + rut + '_dbt_' + i),
        date:   r[2] ? String(r[2]) : '',
        text:   String(r[3] || ''),
        amount: Number(r[4]) || 0
      }));

    // Categorías
    const cats    = categorias.filter(r => String(r[0]).toUpperCase() === rut);
    const catsInc = cats.filter(r => r[1] === 'INC').sort((a,b) => a[3]-b[3]).map(r => String(r[2]));
    const catsExp = cats.filter(r => r[1] === 'EXP').sort((a,b) => a[3]-b[3]).map(r => String(r[2]));
    const catsSav = cats.filter(r => r[1] === 'SAV').sort((a,b) => a[3]-b[3]).map(r => String(r[2]));

    // Archivos
    const arcs = archivos
      .filter(r => String(r[0]).toUpperCase() === rut)
      .map(r => {
        try   { return { id: String(r[1]), period: String(r[2]), data: JSON.parse(String(r[3])) }; }
        catch (_) { return { id: String(r[1]), period: String(r[2]), data: {} }; }
      });

    // ── Diagnósticos ────────────────────────────────────────────────────────
    Logger.log('→ ' + rut + ' (' + nombre + '): '
      + txs.length + ' tx, '
      + userDebts.length + ' deudas, '
      + (catsInc.length + catsExp.length + catsSav.length) + ' cats');

    // Detectar IDs vacíos o duplicados en movimientos
    const ids = txs.map(t => t.id);
    const emptyIds = ids.filter(id => !id || id === 'undefined' || id === '0').length;
    if (emptyIds > 0) Logger.log('  ⚠️ ' + emptyIds + ' movimientos con ID vacío');

    const idSet = new Set(ids);
    if (idSet.size < ids.length) {
      Logger.log('  ⚠️ IDs duplicados en movimientos: ' + (ids.length - idSet.size) + ' duplicados');
      const counts = {};
      ids.forEach(id => { counts[id] = (counts[id] || 0) + 1; });
      const dupes = Object.entries(counts).filter(([_, n]) => n > 1).slice(0, 5);
      Logger.log('  Ejemplos: ' + JSON.stringify(dupes));
    }

    if (ids.length > 0) Logger.log('  Primeros IDs: ' + ids.slice(0, 3).join(' | '));
    // ────────────────────────────────────────────────────────────────────────

    const payload = {
      action: 'migrate',
      secret: MIGRATION_SECRET,
      rut, nombre, password_hash,
      data: { txs, debts: userDebts, archives: arcs, catsInc, catsExp, catsSav }
    };

    try {
      const response = UrlFetchApp.fetch(API_URL, {
        method:             'post',
        contentType:        'application/json',
        payload:            JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const code = response.getResponseCode();
      const text = response.getContentText();

      if (code === 200) {
        try {
          const resp = JSON.parse(text);
          const c = resp.counts || {};
          Logger.log('✓ ' + rut + ' OK — servidor confirmó: '
            + c.txs + ' tx, ' + c.debts + ' deudas, ' + c.cats + ' cats, ' + c.archives + ' archivos');
        } catch (_) {
          Logger.log('✓ ' + rut + ' OK');
        }
        ok++;
      } else {
        Logger.log('✗ ' + rut + ' ERROR ' + code + ': ' + text);
        errors++;
      }
    } catch (e) {
      Logger.log('✗ ' + rut + ' EXCEPCIÓN: ' + e.message);
      errors++;
    }
  }

  Logger.log('\n══════════════════════════════════');
  Logger.log('RESULTADO: ' + ok + ' migrados, ' + errors + ' errores');
  Logger.log('══════════════════════════════════');
}
