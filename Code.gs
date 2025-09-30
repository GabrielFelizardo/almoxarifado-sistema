// =======================================================================
// CONFIGURAÇÃO
// =======================================================================
const SPREADSHEET_ID = "1aOWeZYIo1Che6vlj7TPwasVcxmR97eeq83WsWsDtP_I";
const EMAIL_DESTINO = "gfelizardo14@gmail.com";
const SHEET_NAME_PEDIDOS = "Pedidos";
const SHEET_NAME_ESTOQUE = "Estoque";
const LOW_STOCK_THRESHOLD = 10;
const HIGH_STOCK_THRESHOLD = 100;

function getSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    pedidos: ss.getSheetByName(SHEET_NAME_PEDIDOS),
    estoque: ss.getSheetByName(SHEET_NAME_ESTOQUE)
  };
}

// =======================================================================
// Roteadores Principais (GET e POST)
// =======================================================================
function doGet(e) {
  try {
    const page = e.parameter.page || 'inventory';
    let data;
    if (page === 'inventory') { data = getInventoryData(); }
    else if (page === 'orders') { data = getOrdersData(); }
    else if (page === 'analytics') { data = getAnalyticsData(); }
    else { throw new Error("Página solicitada inválida."); }
    return createJsonResponse(data);
  } catch (error) {
    Logger.log(`GET Error: ${error.message}\n${error.stack}`);
    return createJsonResponse({ error: error.toString() });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 
  try {
    const data = JSON.parse(e.postData.contents);
    let result;
    switch (data.action) {
      case 'submitOrder': result = handleOrderSubmission(data); break;
      case 'addStock': result = handleAddStock(data); break;
      case 'addNewItem': result = handleAddNewItem(data); break;
      case 'updateOrderStatus': result = handleUpdateStatus(data); break;
      case 'deleteItem': result = handleDeleteItem(data); break;
      default: throw new Error("Ação desconhecida.");
    }
    return createJsonResponse(result);
  } catch (error) {
    Logger.log(`POST Error: ${error.message}\n${error.stack}`);
    return createJsonResponse({ result: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

// =======================================================================
// Funções de Leitura de Dados (GET)
// =======================================================================
function getInventoryData() {
  const { estoque } = getSheets();
  if (estoque.getLastRow() < 2) return {};
  const values = estoque.getRange(2, 1, estoque.getLastRow() - 1, 4).getValues();
  const inventoryByCategory = {};
  values.forEach(([code, itemName, quantity, category]) => {
    if (itemName && category) {
      if (!inventoryByCategory[category]) inventoryByCategory[category] = [];
      inventoryByCategory[category].push({ sku: code, nome: itemName, qtd: quantity });
    }
  });
  return inventoryByCategory;
}

function getOrdersData() {
  const { pedidos } = getSheets();
  if (pedidos.getLastRow() < 2) return [];
  const values = pedidos.getRange(2, 1, pedidos.getLastRow() - 1, 9).getValues();
  return values.map(([timestamp, pedidoId, solicitante, setor, idFuncional, item, qtd, justificativa, status]) => ({
    timestamp, pedidoId, solicitante, setor, idFuncional, item, qtd, justificativa, status
  })).reverse();
}

function getAnalyticsData() {
    const { estoque, pedidos } = getSheets();
    const estoqueValues = (estoque.getLastRow() > 1) ? estoque.getRange(2, 2, estoque.getLastRow() - 1, 2).getValues() : [];
    const lowStockItems = estoqueValues.filter(row => Number(row[1]) <= LOW_STOCK_THRESHOLD);
    const highStockItems = estoqueValues.filter(row => Number(row[1]) >= HIGH_STOCK_THRESHOLD);
    const allPedidos = (pedidos.getLastRow() > 1) ? pedidos.getRange(2, 1, pedidos.getLastRow() - 1, 9).getValues() : [];

    const now = new Date();
    const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const dayOfWeek = now.getDay();
    const startOfWeek = new Date(now);
    startOfWeek.setDate(now.getDate() - dayOfWeek);
    startOfWeek.setHours(0, 0, 0, 0);
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);

    const ordersToday = allPedidos.filter(row => row[0] && new Date(row[0]) >= startOfToday);
    const ordersThisWeek = allPedidos.filter(row => row[0] && new Date(row[0]) >= startOfWeek);
    const ordersThisMonth = allPedidos.filter(row => row[0] && new Date(row[0]) >= startOfMonth);
    const pendingOrders = allPedidos.filter(row => row[8] && (row[8] === 'Recebido' || row[8] === 'Em separação'));

    const getTopItems = (orderList) => {
        if (!orderList || orderList.length === 0) return [];
        const counts = orderList.reduce((acc, row) => {
            const item = row[5]; // Coluna F (Item)
            const qty = Number(row[6]); // Coluna G (Quantidade)
            if(item && qty) { acc[item] = (acc[item] || 0) + qty; }
            return acc;
        }, {});
        return Object.entries(counts).sort(([,a],[,b]) => b-a).slice(0, 5).map(([name, count]) => ({name, count}));
    };

    return {
        lowStockItems, highStockItems,
        mostRequestedDay: getTopItems(ordersToday),
        mostRequestedWeek: getTopItems(ordersThisWeek),
        mostRequestedMonth: getTopItems(ordersThisMonth),
        pendingOrdersCount: pendingOrders.length,
        ordersTodayCount: ordersToday.length
    };
}

// =======================================================================
// Funções de Escrita de Dados (POST)
// =======================================================================
function handleOrderSubmission(data) {
  const { pedidos } = getSheets();
  const timestamp = new Date();
  const pedidoId = "PED-" + timestamp.getTime();
  data.itens.forEach(item => {
    // NÃO dá baixa no estoque aqui. Adiciona o ID Funcional.
    pedidos.appendRow([timestamp, pedidoId, data.solicitante, data.setor, data.idFuncional, item.nome, item.qtd, data.justificativa, "Recebido"]);
  });
  enviarEmailNotificacao(data);
  return { result: 'success' };
}

function handleUpdateStatus(data) {
  const { pedidos, estoque } = getSheets();
  if (!data.pedidoId || !data.newStatus) {
    throw new Error("Dados para atualizar status incompletos.");
  }

  const allPedidosRange = pedidos.getRange(2, 1, pedidos.getLastRow() - 1, 9);
  const allPedidosValues = allPedidosRange.getValues();
  let updated = false;

  for (let i = 0; i < allPedidosValues.length; i++) {
    const rowData = allPedidosValues[i];
    const currentPedidoId = rowData[1]; // Coluna B: ID do Pedido

    if (currentPedidoId == data.pedidoId) {
      // Atualiza o status na planilha
      pedidos.getRange(i + 2, 9).setValue(data.newStatus); // Coluna I: Status
      updated = true;

      // Se o novo status é "Concluído" E o status anterior não era "Concluído"
      const oldStatus = rowData[8];
      if (data.newStatus === 'Concluído' && oldStatus !== 'Concluído') {
        const itemName = rowData[5]; // Coluna F: Item
        const itemQty = Number(rowData[6]); // Coluna G: Quantidade
        Logger.log(`Status 'Concluído' para o pedido ${data.pedidoId}, item ${itemName}. Acionando baixa de ${itemQty} unidade(s).`);
        darBaixaEstoque(estoque, itemName, itemQty);
      }
    }
  }

  if (updated) {
    return { result: 'success' };
  } else {
    throw new Error("ID do Pedido não encontrado.");
  }
}

// Funções auxiliares (demais funções permanecem as mesmas)
function handleAddStock(data) { const { estoque } = getSheets(); const values = estoque.getRange(2, 2, estoque.getLastRow() - 1, 2).getValues(); for (let i = 0; i < values.length; i++) { if (values[i][0].toString().trim().toLowerCase() === data.itemName.trim().toLowerCase()) { const estoqueAtual = parseInt(values[i][1]) || 0; const novoEstoque = estoqueAtual + parseInt(data.quantityToAdd); estoque.getRange(i + 2, 3).setValue(novoEstoque); return { result: 'success', newStock: novoEstoque }; } } throw new Error("Item não encontrado."); }
function handleAddNewItem(data) { const { estoque } = getSheets(); const values = estoque.getRange(2, 2, estoque.getLastRow(), 1).getValues(); const newItemNameLower = data.itemName.trim().toLowerCase(); if (values.some(row => row[0].toString().trim().toLowerCase() === newItemNameLower)) { throw new Error(`O item "${data.itemName}" já existe.`); } const newSKU = generateNewSKU(estoque, data.category); estoque.appendRow([newSKU, data.itemName.trim(), data.initialQuantity, data.category]); return { result: 'success' }; }
function handleDeleteItem(data) { const { estoque } = getSheets(); const values = estoque.getRange(2, 2, estoque.getLastRow() - 1, 1).getValues(); const itemNameToDeleteLower = data.itemName.trim().toLowerCase(); for (let i = 0; i < values.length; i++) { if (values[i][0].toString().trim().toLowerCase() === itemNameToDeleteLower) { estoque.deleteRow(i + 2); return { result: 'success' }; } } throw new Error("Item não encontrado para exclusão."); }
function darBaixaEstoque(sheet, nomeItem, quantidadeBaixar) { const values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getValues(); for (let i = 0; i < values.length; i++) { if (values[i][0].toString().trim().toLowerCase() === nomeItem.trim().toLowerCase()) { const estoqueAtual = parseInt(values[i][1]) || 0; if (estoqueAtual >= quantidadeBaixar) { const novoEstoque = estoqueAtual - quantidadeBaixar; sheet.getRange(i + 2, 3).setValue(novoEstoque); if (novoEstoque > 0 && novoEstoque <= LOW_STOCK_THRESHOLD) { enviarAlertaEstoqueBaixo(nomeItem, novoEstoque); } } return; } } }
function generateNewSKU(sheet, category) { const prefixMap = { 'Material de Escritório': 'ESC', 'Material para Copa': 'COP', 'Material de Limpeza': 'LMP' }; const prefix = prefixMap[category] || 'GER'; const codes = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat(); let maxNum = 0; codes.forEach(code => { if (code && code.startsWith(prefix)) { const num = parseInt(code.split('-')[1]); if (num > maxNum) { maxNum = num; } } }); const newNum = maxNum + 1; return `${prefix}-${String(newNum).padStart(3, '0')}`; }
function enviarAlertaEstoqueBaixo(itemName, stockLevel) { const assunto = `Alerta de Estoque Baixo: ${itemName}`; const corpoHtml = `<h2>Atenção: Estoque Baixo</h2><p>O material <strong>${itemName}</strong> atingiu um nível de estoque crítico.</p><p><strong>Quantidade restante:</strong> ${stockLevel}</p><p>Por favor, providencie a reposição do item.</p><p><em>Este é um alerta automático do Sistema de Almoxarifado.</em></p>`; MailApp.sendEmail({ to: EMAIL_DESTINO, subject: assunto, htmlBody: corpoHtml }); }
function enviarEmailNotificacao(data) { const itensHtml = data.itens.map(item => `<li>${item.nome} (Quantidade: ${item.qtd})</li>`).join(''); const assunto = `Novo Pedido de Material - Setor ${data.setor}`; const corpoHtml = `<h2>Novo Pedido de Material Recebido</h2><p><strong>Solicitante:</strong> ${data.solicitante}</p><p><strong>Setor:</strong> ${data.setor}</p><p><strong>ID Funcional:</strong> ${data.idFuncional || 'Não informado'}</p><hr><h4>Itens Solicitados:</h4><ul>${itensHtml}</ul><p><em>Este e-mail foi gerado automaticamente pelo Sistema de Almoxarifado.</em></p>`; MailApp.sendEmail({ to: EMAIL_DESTINO, subject: assunto, htmlBody: corpoHtml }); }
function createJsonResponse(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
