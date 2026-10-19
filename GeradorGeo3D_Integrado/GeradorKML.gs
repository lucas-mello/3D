
/**
 * Gerador de polígonos KML — GeradorKML.gs
 *
 * Integração: adicione .addItem('Gerar polígonos KML', 'abrirGeradorKML')
 * ao onOpen() já existente. Não crie outro onOpen() no projeto.
 *
 * Planilha:
 * A = número do ponto
 * C = Norte (Y)
 * D = Este (X)
 * E = altitude
 *
 * O HTML pode enviar:
 * fechado: true  -> polígono fechado
 * fechado: false -> linha aberta, sem segmento de fechamento
 *
 * Se "fechado" não for informado, será considerado true.
 */

function abrirGeradorKML() {
  const html = HtmlService.createHtmlOutputFromFile('GeradorkML')
    .setWidth(1100)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, 'Gerar polígonos KML');
}


function kmlObterConfiguracaoInicial() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const lastRow = sh.getLastRow();

  if (lastRow < 2) {
    return {
      sheetName: sh.getName(),
      points: [],
      message: 'A aba ativa não contém pontos.'
    };
  }

  const rows = sh.getRange(2, 1, lastRow - 1, 5).getValues();
  const points = [];

  rows.forEach((r, idx) => {
    const num = String(r[0] || '').trim();
    const north = kmlNumero(r[2]);
    const east = kmlNumero(r[3]);
    const alt = kmlNumero(r[4]);

    if (num && isFinite(north) && isFinite(east)) {
      points.push({
        number: num,
        north: north,
        east: east,
        altitude: isFinite(alt) ? alt : 0,
        row: idx + 2
      });
    }
  });

  points.forEach(p => {
    const ll = kmlUtmParaLonLat(p.east, p.north, 23, false);
    p.lon = ll.lon;
    p.lat = ll.lat;
  });

  return {
    sheetName: sh.getName(),
    points: points,
    message: points.length
      ? ''
      : 'Não encontrei pontos válidos. Confira as colunas A, C, D e E.'
  };
}


function kmlNumero(value) {
  if (typeof value === 'number') return value;

  let s = String(value == null ? '' : value)
    .trim()
    .replace(/\s/g, '');

  if (!s) return NaN;

  if (s.includes(',') && s.includes('.')) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      s = s.replace(/,/g, '');
    }
  } else if (s.includes(',')) {
    s = s.replace(',', '.');
  }

  return Number(s);
}


/**
 * Salva os pontos do grupo na planilha e devolve as coordenadas geográficas.
 * Grupos abertos podem ter 2 ou mais pontos.
 * Grupos fechados precisam ter pelo menos 3 pontos.
 */
function kmlSalvarGrupoEObterCoordenadas(
  numeros,
  grupoIndex,
  zona,
  hemisferio,
  fechado
) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();

  // Mantém o comportamento antigo quando o parâmetro não é informado.
  fechado = fechado !== false;

  const minimo = fechado ? 3 : 2;

  if (!Array.isArray(numeros) || numeros.length < minimo) {
    if (fechado) {
      throw new Error(
        'Informe pelo menos 3 pontos para formar um polígono fechado.'
      );
    }

    throw new Error(
      'Informe pelo menos 2 pontos para formar uma linha aberta.'
    );
  }

  const cfg = kmlObterConfiguracaoInicial();
  const byNum = {};

  cfg.points.forEach(p => {
    byNum[String(Number(p.number))] = p;
    byNum[p.number] = p;
  });

  const chosen = numeros.map(n => {
    const key = String(n).trim();
    const p = byNum[key] || byNum[String(Number(key))];

    if (!p) {
      throw new Error(
        'O ponto ' + key + ' não foi encontrado na coluna A.'
      );
    }

    return p;
  });

  // Grava os pontos a partir da coluna I.
  const col = 9 + Number(grupoIndex);
  const tipo = fechado ? 'POLÍGONO' : 'LINHA ABERTA';

  const title = tipo + ' ' + (Number(grupoIndex) + 1) +
    ' — PONTOS: ' + numeros.join(', ');

  sh.getRange(1, col).setValue(title);

  const out = chosen.map(p => [
    p.number + ' | ' + p.east + ',' + p.north +
    (p.altitude ? ',' + p.altitude : '')
  ]);

  const rowsToClear = Math.max(sh.getMaxRows() - 1, out.length);

  if (rowsToClear > 0) {
    sh.getRange(2, col, rowsToClear, 1).clearContent();
  }

  sh.getRange(2, col, out.length, 1)
    .setNumberFormat('@')
    .setValues(out);

  return {
    name: title,
    fechado: fechado,
    points: chosen.map(p => {
      const ll = kmlUtmParaLonLat(
        p.east,
        p.north,
        Number(zona),
        String(hemisferio).toUpperCase() === 'N'
      );

      return {
        number: p.number,
        lon: ll.lon,
        lat: ll.lat,
        altitude: p.altitude
      };
    })
  };
}


/**
 * Gera o arquivo KML.
 *
 * fechado = true:
 *   Gera Polygon e fecha o anel de coordenadas.
 *
 * fechado = false:
 *   Gera LineString e NÃO repete o primeiro ponto no final.
 */
function kmlGerarArquivo(grupos) {
  if (!Array.isArray(grupos) || !grupos.length) {
    throw new Error('Nenhum grupo foi adicionado.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();

  const zona = Number(grupos[0].zona);
  const norte = String(grupos[0].hemisferio).toUpperCase() === 'N';

  let xml =
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
    '<kml xmlns="http://www.opengis.net/kml/2.2">' +
    '<Document>' +
    '<name>' + kmlEscapar(sh.getName()) + '</name>\n';

  grupos.forEach((g, i) => {
    // Se o HTML não enviar "fechado", preserva o comportamento antigo.
    const fechado = g.fechado !== false;

    const data = kmlSalvarGrupoEObterCoordenadas(
      g.numeros,
      i,
      zona,
      norte ? 'N' : 'S',
      fechado
    );

    const coords = data.points
      .map(p =>
        p.lon.toFixed(8) + ',' +
        p.lat.toFixed(8) + ',' +
        (p.altitude || 0)
      )
      .join(' ');

    let coordenadasKml = coords;

    // Só repete o primeiro ponto quando o grupo é fechado.
    if (fechado) {
      const first = data.points[0];

      coordenadasKml += ' ' +
        first.lon.toFixed(8) + ',' +
        first.lat.toFixed(8) + ',' +
        (first.altitude || 0);
    }

    const nm = g.name || data.name;
    const desc = g.description ||
      ('Pontos: ' + g.numeros.join(', '));

    const color = String(g.color || 'ff00a5ff').replace('#', '');
    const kmlColor = color.length === 8 ? color : 'ff00a5ff';

    xml +=
      '<Style id="style' + i + '">' +
        '<LineStyle>' +
          '<color>' + kmlColor + '</color>' +
          '<width>2</width>' +
        '</LineStyle>' +
        '<PolyStyle>' +
          '<color>5500a5ff</color>' +
          '<fill>' + (fechado ? '1' : '0') + '</fill>' +
          '<outline>1</outline>' +
        '</PolyStyle>' +
      '</Style>';

    xml +=
      '<Placemark>' +
        '<name>' + kmlEscapar(nm) + '</name>' +
        '<description>' + kmlEscapar(desc) + '</description>' +
        '<styleUrl>#style' + i + '</styleUrl>';

    if (fechado) {
      // Polígono fechado.
      xml +=
        '<Polygon>' +
          '<altitudeMode>clampToGround</altitudeMode>' +
          '<outerBoundaryIs>' +
            '<LinearRing>' +
              '<coordinates>' + coordenadasKml + '</coordinates>' +
            '</LinearRing>' +
          '</outerBoundaryIs>' +
        '</Polygon>';
    } else {
      // Linha aberta: não conecta o último ponto ao primeiro.
      xml +=
        '<LineString>' +
          '<altitudeMode>clampToGround</altitudeMode>' +
          '<tessellate>1</tessellate>' +
          '<coordinates>' + coordenadasKml + '</coordinates>' +
        '</LineString>';
    }

    xml += '</Placemark>\n';
  });

  xml += '</Document></kml>';

  return {
    filename: sh.getName().replace(/[^a-zA-Z0-9_-]/g, '_') +
      '_poligonos.kml',
    content: xml
  };
}


function kmlEscapar(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}


/**
 * Conversão UTM para longitude/latitude.
 * SIRGAS 2000 / UTM zona 23S é usado como referência para Belo Horizonte.
 * O KML exige coordenadas na ordem longitude, latitude, altitude.
 */
function kmlUtmParaLonLat(easting, northing, zone, northHemisphere) {
  const a = 6378137.0;
  const eccSquared = 0.00669438002290;
  const k0 = 0.9996;
  const eccPrimeSquared = eccSquared / (1 - eccSquared);

  const x = easting - 500000.0;
  let y = northing;

  if (!northHemisphere) {
    y -= 10000000.0;
  }

  const longOrigin = (zone - 1) * 6 - 180 + 3;
  const M = y / k0;

  const mu = M / (
    a * (
      1 -
      eccSquared / 4 -
      3 * eccSquared * eccSquared / 64 -
      5 * eccSquared * eccSquared * eccSquared / 256
    )
  );

  const e1 = (1 - Math.sqrt(1 - eccSquared)) /
    (1 + Math.sqrt(1 - eccSquared));

  const J1 = 3 * e1 / 2 - 27 * Math.pow(e1, 3) / 32;
  const J2 = 21 * e1 * e1 / 16 - 55 * Math.pow(e1, 4) / 32;
  const J3 = 151 * Math.pow(e1, 3) / 96;
  const J4 = 1097 * Math.pow(e1, 4) / 512;

  const fp =
    mu +
    J1 * Math.sin(2 * mu) +
    J2 * Math.sin(4 * mu) +
    J3 * Math.sin(6 * mu) +
    J4 * Math.sin(8 * mu);

  const C1 = eccPrimeSquared * Math.pow(Math.cos(fp), 2);
  const T1 = Math.pow(Math.tan(fp), 2);
  const N1 = a / Math.sqrt(
    1 - eccSquared * Math.pow(Math.sin(fp), 2)
  );

  const R1 = a * (1 - eccSquared) /
    Math.pow(
      1 - eccSquared * Math.pow(Math.sin(fp), 2),
      1.5
    );

  const D = x / (N1 * k0);

  let lat = fp - (N1 * Math.tan(fp) / R1) * (
    D * D / 2 -
    (5 + 3 * T1 + 10 * C1 - 4 * C1 * C1 -
      9 * eccPrimeSquared) * Math.pow(D, 4) / 24 +
    (61 + 90 * T1 + 298 * C1 + 45 * T1 * T1 -
      252 * eccPrimeSquared - 3 * C1 * C1) *
      Math.pow(D, 6) / 720
  );

  lat = lat * 180 / Math.PI;

  let lon = (
    D -
    (1 + 2 * T1 + C1) * Math.pow(D, 3) / 6 +
    (5 - 2 * C1 + 28 * T1 - 3 * C1 * C1 +
      8 * eccPrimeSquared + 24 * T1 * T1) *
      Math.pow(D, 5) / 120
  ) / Math.cos(fp);

  lon = longOrigin + lon * 180 / Math.PI;

  return {
    lon: lon,
    lat: lat
  };
}
