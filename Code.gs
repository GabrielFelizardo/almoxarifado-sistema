// =======================================================================
// CONFIGURAÇÃO
// =======================================================================
const SPREADSHEET_ID = "1aOWeZYIo1Che6vlj7TPwasVcxmR97eeq83WsWsDtP_I";
const EMAIL_DESTINO = "gfelizardo14@gmail.com";
const SHEET_NAME_PEDIDOS = "Pedidos";
const SHEET_NAME_ESTOQUE = "Estoque";
const LOW_STOCK_THRESHOLD = 10;

// Função para acessar a planilha e as abas de forma centralizada
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
    if (page === 'inventory') {
      data = getInventoryData();
    } else if (page === 'orders') {
      data = getOrdersData();
    } else if (page === 'analytics') {
      // Passa todos os possíveis parâmetros de filtro para a função de análise
      data = getAnalyticsData(e.parameter.startDate, e.parameter.endDate, e.parameter.setor, e.parameter.status);
    } else {
      throw new Error("Página solicitada inválida.");
    }
    return createJsonResponse(data);
  } catch (error) {
    Logger.log(`GET Error: ${error.message} \n ${error.stack}`);
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
    Logger.log(`POST Error: ${error.message} \n ${error.stack}`);
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
  const values = pedidos.getRange(2, 1, pedidos.getLastRow() - 1, 8).getValues();
  return values.map(([timestamp, pedidoId, solicitante, setor, item, qtd, justificativa, status]) => ({
    timestamp, pedidoId, solicitante, setor, item, qtd, justificativa, status
  })).reverse();
}

function getAnalyticsData(startDateStr, endDateStr, setor, status) {
    const { estoque, pedidos } = getSheets();

    const estoqueValues = (estoque.getLastRow() > 1) ? estoque.getRange(2, 2, estoque.getLastRow() - 1, 2).getValues() : [];
    const lowStockItems = estoqueValues.filter(row => row[1] <= LOW_STOCK_THRESHOLD);
    
    let allPedidos = (pedidos.getLastRow() > 1) ? pedidos.getRange(2, 1, pedidos.getLastRow() - 1, 8).getValues() : [];

    let filteredPedidos = allPedidos;

    if (startDateStr && endDateStr) {
        const startDate = new Date(startDateStr);
        const endDate = new Date(endDateStr);
        endDate.setHours(23, 59, 59, 999);

        filteredPedidos = filteredPedidos.filter(row => {
            const orderDate = new Date(row[0]);
            return orderDate >= startDate && orderDate <= endDate;
        });
    }

    if (setor) {
        filteredPedidos = filteredPedidos.filter(row => row[3] === setor);
    }

    if (status) {
        filteredPedidos = filteredPedidos.filter(row => row[7] === status);
    }
    
    const requestCounts = filteredPedidos.reduce((acc, row) => {
        const item = row[4];
        const qty = Number(row[5]);
        if(item && qty) {
            acc[item] = (acc[item] || 0) + qty;
        }
        return acc;
    }, {});

    const mostRequested = Object.entries(requestCounts)
        .sort(([, a], [, b]) => b - a)
        .slice(0, 10)
        .map(([name, count]) => ({ name, count }));

    return {
        totalItems: estoqueValues.length,
        lowStockCount: lowStockItems.length,
        lowStockItems: lowStockItems,
        mostRequested: mostRequested,
        period: { start: startDateStr, end: endDateStr },
        filters: { setor: setor, status: status }
    };
}

// =======================================================================
// Funções de Escrita de Dados (POST)
// =======================================================================
function handleOrderSubmission(data) {
  const { pedidos, estoque } = getSheets();
  const timestamp = new Date();
  const pedidoId = "PED-" + timestamp.getTime();
  data.itens.forEach(item => {
    pedidos.appendRow([timestamp, pedidoId, data.solicitante, data.setor, item.nome, item.qtd, data.justificativa, "Recebido"]);
    darBaixaEstoque(estoque, item.nome, Number(item.qtd));
  });
  enviarEmailNotificacao(data);
  return { result: 'success' };
}

function handleAddStock(data) {
  const { estoque } = getSheets();
  const values = estoque.getRange(2, 2, estoque.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0].toString().trim().toLowerCase() === data.itemName.trim().toLowerCase()) {
      const estoqueAtual = parseInt(values[i][1]) || 0;
      const novoEstoque = estoqueAtual + parseInt(data.quantityToAdd);
      estoque.getRange(i + 2, 3).setValue(novoEstoque);
      return { result: 'success', newStock: novoEstoque };
    }
  }
  throw new Error("Item não encontrado.");
}

function handleAddNewItem(data) {
  const { estoque } = getSheets();
  const values = estoque.getRange(2, 2, estoque.getLastRow(), 1).getValues();
  const newItemNameLower = data.itemName.trim().toLowerCase();
  if (values.some(row => row[0].toString().trim().toLowerCase() === newItemNameLower)) {
    throw new Error(`O item "${data.itemName}" já existe.`);
  }
  const newSKU = generateNewSKU(estoque, data.category);
  estoque.appendRow([newSKU, data.itemName.trim(), data.initialQuantity, data.category]);
  return { result: 'success' };
}

function handleUpdateStatus(data) {
  const { pedidos } = getSheets();
  const pedidoIds = pedidos.getRange(2, 2, pedidos.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < pedidoIds.length; i++) {
    if (pedidoIds[i][0] == data.pedidoId) {
      pedidos.getRange(i + 2, 8).setValue(data.newStatus);
      return { result: 'success' };
    }
  }
  throw new Error("ID do Pedido não encontrado.");
}

function handleDeleteItem(data) {
  const { estoque } = getSheets();
  const values = estoque.getRange(2, 2, estoque.getLastRow() - 1, 1).getValues();
  const itemNameToDeleteLower = data.itemName.trim().toLowerCase();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0].toString().trim().toLowerCase() === itemNameToDeleteLower) {
      estoque.deleteRow(i + 2);
      return { result: 'success' };
    }
  }
  throw new Error("Item não encontrado para exclusão.");
}

// =======================================================================
// Funções Auxiliares
// =======================================================================
function darBaixaEstoque(sheet, nomeItem, quantidadeBaixar) {
  const values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0].toString().trim().toLowerCase() === nomeItem.trim().toLowerCase()) {
      const estoqueAtual = parseInt(values[i][1]) || 0;
      if (estoqueAtual >= quantidadeBaixar) {
        const novoEstoque = estoqueAtual - quantidadeBaixar;
        sheet.getRange(i + 2, 3).setValue(novoEstoque);
        if (novoEstoque > 0 && novoEstoque <= LOW_STOCK_THRESHOLD) {
          enviarAlertaEstoqueBaixo(nomeItem, novoEstoque);
        }
      }
      return;
    }
  }
}

function generateNewSKU(sheet, category) {
    const prefixMap = { 'Material de Escritório': 'ESC', 'Material para Copa': 'COP', 'Material de Limpeza': 'LMP' };
    const prefix = prefixMap[category] || 'GER';
    const codes = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
    let maxNum = 0;
    codes.forEach(code => {
        if (code && code.startsWith(prefix)) {
            const num = parseInt(code.split('-')[1]);
            if (num > maxNum) { maxNum = num; }
        }
    });
    const newNum = maxNum + 1;
    return `${prefix}-${String(newNum).padStart(3, '0')}`;
}

function enviarAlertaEstoqueBaixo(itemName, stockLevel) {
    const assunto = `Alerta de Estoque Baixo: ${itemName}`;
    const corpoHtml = `<h2>Atenção: Estoque Baixo</h2><p>O material <strong>${itemName}</strong> atingiu um nível de estoque crítico.</p><p><strong>Quantidade restante:</strong> ${stockLevel}</p><p>Por favor, providencie a reposição do item.</p><p><em>Este é um alerta automático do Sistema de Almoxarifado.</em></p>`;
    MailApp.sendEmail({ to: EMAIL_DESTINO, subject: assunto, htmlBody: corpoHtml });
}

function enviarEmailNotificacao(data) {
    const itensHtml = data.itens.map(item => `<li>${item.nome} (Quantidade: ${item.qtd})</li>`).join('');
    const assunto = `Novo Pedido de Material - Setor ${data.setor}`;
    const corpoHtml = `<h2>Novo Pedido de Material Recebido</h2><p><strong>Solicitante:</strong> ${data.solicitante}</p><p><strong>Setor:</strong> ${data.setor}</p><p><strong>Justificativa:</strong> ${data.justificativa || 'Nenhuma'}</p><hr><h4>Itens Solicitados:</h4><ul>${itensHtml}</ul><p><em>Este e-mail foi gerado automaticamente pelo Sistema de Almoxarifado.</em></p>`;
    MailApp.sendEmail({ to: EMAIL_DESTINO, subject: assunto, htmlBody: corpoHtml });
}

function createJsonResponse(data) {
    return ContentService.createTextOutput(JSON.stringify(data))
                         .setMimeType(ContentService.MimeType.JSON);
}
