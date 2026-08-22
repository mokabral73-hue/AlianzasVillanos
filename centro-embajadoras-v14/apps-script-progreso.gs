const SHEET_ID = '1PCbiXpgbSdyx28wbqtkjSu0Dz4uuSJwQZH4pEnGBL6c';
const SHEET_NAME = 'Progreso';

function doPost(e) {
  try {
    const p = e.parameter || {};
    const id = String(p.id || '').trim();
    if (!id) return json_({ ok:false, error:'missing_id' });

    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (!sh) return json_({ ok:false, error:'sheet_not_found' });

    const lastRow = Math.max(sh.getLastRow(), 1);
    const ids = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
    let row = ids.findIndex(v => String(v) === id);
    row = row >= 0 ? row + 2 : lastRow + 1;

    const now = new Date();
    const values = [[
      id,
      p.nombre || '',
      p.ciudad || '',
      p.nivel || '',
      p.estacion || '',
      p.estado || 'Activa',
      Number(p.acompanadas || 0),
      Number(p.invitadas || 0),
      Number(p.nuevas || 0),
      now
    ]];

    sh.getRange(row, 1, 1, 10).setValues(values);
    return json_({ ok:true, row:row });
  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

function doGet() {
  return json_({ ok:true, service:'Menopausia Dichosa · Progreso Embajadoras' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
