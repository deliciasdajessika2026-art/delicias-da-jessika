/**
 * ============================================================
 * Delícias da Jéssika — Backend (Google Apps Script)
 * Banco de dados real usando Google Sheets + estrutura no Google Drive
 * ============================================================
 *
 * COMO PUBLICAR (faça isso uma única vez):
 * 1. Acesse https://script.google.com e crie um novo projeto.
 * 2. Apague o conteúdo padrão de "Code.gs" e cole todo o conteúdo deste arquivo.
 * 3. Troque o valor de API_KEY abaixo por uma chave secreta só sua
 *    (qualquer texto longo e difícil de adivinhar).
 * 4. Clique em "Implantar" > "Nova implantação".
 *    - Tipo: "App da Web"
 *    - Executar como: "Eu" (sua conta)
 *    - Quem tem acesso: "Qualquer pessoa"
 *      (isso é necessário para o navegador conseguir chamar a API sem
 *       precisar de login OAuth próprio; a proteção real é a API_KEY)
 * 5. Autorize as permissões pedidas (Planilhas e Drive da sua própria conta).
 * 6. Copie a URL gerada (termina em /exec) e cole em:
 *    Configurações > Integração — Google Sheets, no sistema Delícias da Jéssika.
 * 7. Cole a MESMA chave que você definiu em API_KEY no campo "Chave de API".
 *
 * O que este script faz:
 * - Cria automaticamente (se não existir) uma planilha chamada
 *   "Delícias da Jéssika - Banco de Dados" dentro de uma pasta do Drive
 *   chamada "Delícias da Jéssika" (com subpastas: Backups, Financeiro,
 *   Estoque, Clientes, Ingredientes, Produtos, Produção, Relatórios,
 *   Imagens, Documentos).
 * - Cria uma aba por coleção de dados (Ingredientes, Produtos, Vendas etc.).
 * - Cada aba guarda o conteúdo da coleção como um bloco JSON (célula B1),
 *   por simplicidade e para não exigir um mapeamento rígido de colunas
 *   para cada uma das ~14 entidades do sistema. Isso ainda cumpre o
 *   objetivo de ter os dados persistidos de verdade no Google Sheets,
 *   com sincronização automática a partir do navegador.
 */

var API_KEY = 'TROQUE_ESTA_CHAVE_POR_UMA_SENHA_SECRETA_SUA';
var SPREADSHEET_NAME = 'Delícias da Jéssika - Banco de Dados';
var DRIVE_FOLDER_NAME = 'Delícias da Jéssika';

var COLLECTIONS = [
  'config', 'sabores', 'categoriasProdutos', 'ingredientes', 'produtos',
  'producao', 'estoqueProdutos', 'movimentacoes', 'perdas', 'vendas',
  'clientes', 'financeiro', 'auditLog', 'usuarios'
  // 'usuarios' guarda apenas o hash da senha (nunca a senha em texto puro),
  // sincronizado para permitir login com a mesma conta em qualquer aparelho.
];

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.apiKey !== API_KEY) {
      return jsonResponse({ ok: false, error: 'Chave de API inválida' });
    }
    var action = body.action;
    if (action === 'ping') return jsonResponse({ ok: true, message: 'pong' });
    if (action === 'saveCollection') return saveCollection(body.collection, body.data);
    if (action === 'getCollection') return getCollection(body.collection);
    if (action === 'getAll') return getAll();
    return jsonResponse({ ok: false, error: 'Ação desconhecida: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  return jsonResponse({ ok: true, message: 'Delícias da Jéssika API ativa. Use requisições POST.' });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateFolder() {
  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  var root = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
  var subpastas = ['Backups', 'Financeiro', 'Estoque', 'Clientes', 'Ingredientes', 'Produtos', 'Produção', 'Relatórios', 'Imagens', 'Documentos'];
  for (var i = 0; i < subpastas.length; i++) {
    if (!root.getFoldersByName(subpastas[i]).hasNext()) {
      root.createFolder(subpastas[i]);
    }
  }
  return root;
}

function getOrCreateSpreadsheet() {
  var folder = getOrCreateFolder();
  var files = folder.getFilesByName(SPREADSHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  var ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  var file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  return ss;
}

function getOrCreateSheet(ss, nome) {
  var sheet = ss.getSheetByName(nome);
  if (!sheet) {
    sheet = ss.insertSheet(nome);
    sheet.getRange(1, 1).setValue('chave');
    sheet.getRange(1, 2).setValue('dados_json');
    sheet.getRange(1, 3).setValue('atualizado_em');
  }
  return sheet;
}

function saveCollection(nome, data) {
  if (COLLECTIONS.indexOf(nome) === -1) {
    return jsonResponse({ ok: false, error: 'Coleção não permitida: ' + nome });
  }
  var ss = getOrCreateSpreadsheet();
  var sheet = getOrCreateSheet(ss, nome);
  sheet.getRange(2, 1).setValue(nome);
  sheet.getRange(2, 2).setValue(JSON.stringify(data));
  sheet.getRange(2, 3).setValue(new Date().toISOString());
  return jsonResponse({ ok: true, collection: nome, savedAt: new Date().toISOString() });
}

function getCollection(nome) {
  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName(nome);
  if (!sheet) return jsonResponse({ ok: true, collection: nome, data: null });
  var json = sheet.getRange(2, 2).getValue();
  var data = null;
  try { data = json ? JSON.parse(json) : null; } catch (err) { data = null; }
  return jsonResponse({ ok: true, collection: nome, data: data });
}

function getAll() {
  var ss = getOrCreateSpreadsheet();
  var out = {};
  for (var i = 0; i < COLLECTIONS.length; i++) {
    var nome = COLLECTIONS[i];
    var sheet = ss.getSheetByName(nome);
    if (sheet) {
      var json = sheet.getRange(2, 2).getValue();
      try { out[nome] = json ? JSON.parse(json) : null; } catch (err) { out[nome] = null; }
    } else {
      out[nome] = null;
    }
  }
  return jsonResponse({ ok: true, data: out });
}
