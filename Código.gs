const FOLDER_ID = '1_5M1q3uoLPUtsTHr1ht8vJOyDi5DYUsz'; // ID de tu carpeta de Drive

// Clientes externos: se muestran en el apartado "Clientes Externos", separados de los Holding's.
// Se comparan sin mayúsculas, acentos ni espacios ("Apple Town" = "Appletown").
const CLIENTES_EXTERNOS = ['Oriental', 'Appletown'];

// Si cambias la lista de arriba, sube este número para que no se use la caché vieja.
const VERSION_CACHE = 'v3';

function doGet() {
  /* createHtmlOutputFromFile: entrega el HTML tal cual, sin pasar por el motor de plantillas */
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Dashboard IUNGO - Beneficio Médico')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getExcelFiles() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.searchFiles('mimeType="application/vnd.google-apps.spreadsheet"');
  let fileList = [];

  while (files.hasNext()) {
    let file = files.next();
    fileList.push({
      id: file.getId(),
      name: file.getName(),
      updated: file.getLastUpdated().getTime() // número, no Date
    });
  }

  const meses = {
    'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6,
    'julio': 7, 'agosto': 8, 'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12
  };

  const obtenerFecha = (nombre) => {
    let n = nombre.toLowerCase();
    let mesValor = 0, anioValor = 0;
    for (let mes in meses) {
      if (n.includes(mes)) { mesValor = meses[mes]; break; }
    }
    let matchAnio = n.match(/\b(20\d{2})\b/);
    if (matchAnio) anioValor = parseInt(matchAnio[1], 10);
    return { mes: mesValor, anio: anioValor };
  };

  fileList.sort((a, b) => {
    let fa = obtenerFecha(a.name), fb = obtenerFecha(b.name);
    if (fa.anio !== fb.anio) return fb.anio - fa.anio;
    return fb.mes - fa.mes;
  });

  return fileList;
}

// Convierte cualquier valor de celda a algo que el navegador pueda recibir.
// (Fechas u otros objetos pueden hacer que google.script.run no responda.)
function limpiarValor(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return v;
  return String(v);
}

function normalizar(texto) {
  return String(texto).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '');
}

function esClienteExterno(empresa) {
  let n = normalizar(empresa);
  return CLIENTES_EXTERNOS.some(ex => n.includes(normalizar(ex)));
}

function aNumero(v) {
  if (typeof v === 'number') return v;
  let n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function getSummaryData(fileId) {
  let ss = SpreadsheetApp.openById(fileId);
  let sheet = ss.getSheets()[0];
  let data = sheet.getDataRange().getValues();

  let summary = [];
  let startIndex = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i].includes("Unidad de Negocio")) { startIndex = i; break; }
  }

  if (startIndex !== -1) {
    let headerRow = data[startIndex];
    let colEmpresa = headerRow.indexOf("Unidad de Negocio");
    let colBeneficiarios = headerRow.indexOf("Beneficiarios + Chequeo Médico");
    let colIVA = headerRow.indexOf("IVA");
    let colTotal = headerRow.indexOf("Total Factura");

    for (let i = startIndex + 1; i < data.length; i++) {
      let empresa = data[i][colEmpresa];
      if (!empresa || empresa === "") continue;

      // La fila "Totales" del Excel mezcla ambos grupos: el dashboard recalcula los totales por grupo
      if (String(empresa).toLowerCase().includes('totales')) continue;

      summary.push({
        empresa: limpiarValor(empresa),
        grupo: esClienteExterno(empresa) ? 'externo' : 'holding',
        beneficiarios: limpiarValor(data[i][colBeneficiarios]),
        iva: limpiarValor(data[i][colIVA]),
        total: limpiarValor(data[i][colTotal])
      });
    }
  }
  return summary;
}

// forzar = true cuando se presiona "Actualizar": vuelve a leer todos los archivos.
// Si no, solo vuelve a abrir los archivos que cambiaron desde la última lectura.
function getAllMonthsData(forzar) {
  let files = getExcelFiles();
  let cache = CacheService.getScriptCache();
  let cacheData = {};

  files.forEach(f => {
    let key = 'res_' + VERSION_CACHE + '_' + f.id + '_' + f.updated;
    let guardado = forzar ? null : cache.get(key);
    if (guardado) {
      cacheData[f.id] = JSON.parse(guardado);
    } else {
      let resumen = getSummaryData(f.id);
      cacheData[f.id] = resumen;
      try { cache.put(key, JSON.stringify(resumen), 21600); } catch (e) { /* si no cabe en caché, se ignora */ }
    }
  });

  return {
    files: files.map(f => ({ id: f.id, name: f.name })),
    data: cacheData
  };
}
