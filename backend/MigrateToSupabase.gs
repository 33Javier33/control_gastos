/**
 * SCRIPT DE MIGRACIÓN — Google Sheets → Supabase
 *
 * INSTRUCCIONES:
 * 1. Abre script.google.com → abre el proyecto de tu planilla de Control Financiero
 * 2. Crea un nuevo archivo (Archivo > Nuevo > Script) y pega este código
 * 3. Selecciona la función "migrateToSupabase" en el menú desplegable
 * 4. Haz clic en ▶ Ejecutar
 * 5. Acepta los permisos cuando te los pida
 * 6. Revisa el Log de Ejecución (Ver > Registros) para ver el resultado
 *
 * Este script puede ejecutarse múltiples veces sin problema (no duplica datos).
 */

function migrateToSupabase() {
  const API_URL         = 'https://lpulmjzboogixbdxxayo.supabase.co/functions/v1/api';
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

    // Movimientos del usuario
    const txs = movimientos
      .filter(r => String(r[0]).toUpperCase() === rut)
      .map(r => ({
        id:       String(r[1]),
        date:     r[2] ? String(r[2]) : '',
        text:     String(r[3] || ''),
        amount:   Number(r[4]) || 0,
        cat:      String(r[5] || ''),
        source:   String(r[6] || ''),
        isSaving: String(r[7]) === 'SI'
      }));

    // Deudas del usuario
    const userDebts = deudas
      .filter(r => String(r[0]).toUpperCase() === rut)
      .map(r => ({
        id:     String(r[1]),
        date:   r[2] ? String(r[2]) : '',
        text:   String(r[3] || ''),
        amount: Number(r[4]) || 0
      }));

    // Categorías del usuario
    const cats    = categorias.filter(r => String(r[0]).toUpperCase() === rut);
    const catsInc = cats.filter(r => r[1] === 'INC').sort((a,b) => a[3]-b[3]).map(r => String(r[2]));
    const catsExp = cats.filter(r => r[1] === 'EXP').sort((a,b) => a[3]-b[3]).map(r => String(r[2]));
    const catsSav = cats.filter(r => r[1] === 'SAV').sort((a,b) => a[3]-b[3]).map(r => String(r[2]));

    // Archivos del usuario
    const arcs = archivos
      .filter(r => String(r[0]).toUpperCase() === rut)
      .map(r => {
        try   { return { id: String(r[1]), period: String(r[2]), data: JSON.parse(String(r[3])) }; }
        catch (_) { return { id: String(r[1]), period: String(r[2]), data: {} }; }
      });

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
        Logger.log('✓ ' + rut + ' (' + nombre + ') — ' + txs.length + ' movimientos, ' + userDebts.length + ' deudas');
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

  SpreadsheetApp.getUi().alert(
    'Migración completada\n\n' +
    '✓ ' + ok + ' usuarios migrados correctamente\n' +
    (errors > 0 ? '✗ ' + errors + ' errores (ver Log de Ejecución)\n' : '') +
    '\nYa puedes usar la app con Supabase.'
  );
}
