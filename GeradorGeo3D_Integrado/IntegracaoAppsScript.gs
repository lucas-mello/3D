/**
 * INTEGRAÇÃO GEO 3D + GERADOR KML
 *
 * 1. Copie as funções deste arquivo para o projeto Apps Script da planilha.
 * 2. Mantenha o onOpen() que já existe no seu projeto; NÃO crie outro.
 * 3. Dentro do onOpen() existente, acrescente:
 *
 *    .addItem('Abrir mapas Geo 3D', 'abrirGeo3D')
 *    .addItem('Gerar polígonos KML', 'abrirGeradorKML')
 *
 * 4. Substitua GEO3D_URL pela URL HTTPS onde a pasta do Geo 3D foi publicada.
 *
 * Os arquivos GeradorKML.gs e GeradorkML.html devem ser adicionados ao Apps Script
 * com esses nomes. O Geo 3D é uma página externa; o gerador de polígonos continua
 * funcionando na planilha em sua janela própria.
 */
const GEO3D_URL = 'https://SUBSTITUA-PELA-URL-PUBLICADA/index.html';

function abrirGeo3D() {
  const html = HtmlService.createHtmlOutput(
    '<!doctype html><html><head><base target="_blank"></head>' +
    '<body style="font:14px Arial;padding:18px">' +
    '<p>O Geo 3D será aberto em uma nova aba.</p>' +
    '<p><a href="' + GEO3D_URL + '" target="_blank" rel="noopener">Abrir Geo 3D — Belo Horizonte</a></p>' +
    '</body></html>'
  ).setWidth(420).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Mapas Geo 3D');
}
