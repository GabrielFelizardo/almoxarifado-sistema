// =======================================================================
// CONFIGURAÇÃO SEGURA - Lendo as Propriedades do Script
// =======================================================================
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const SPREADSHEET_ID = SCRIPT_PROPERTIES.getProperty('SPREADSHEET_ID');
const EMAIL_DESTINO = SCRIPT_PROPERTIES.getProperty('EMAIL_DESTINO');
// =======================================================================

const SHEET_NAME_PEDIDOS = "Pedidos";
const SHEET_NAME_ESTOQUE = "Estoque";
const SHEET_NAME_COMPRAS = "Compras";
const LOW_STOCK_THRESHOLD = 10;
const HIGH_STOCK_THRESHOLD = 100;

function getSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    pedidos: ss.getSheetByName(SHEET_NAME_PEDIDOS),
    estoque: ss.getSheetByName(SHEET_NAME_ESTOQUE),
    compras: ss.getSheetByName(SHEET_NAME_COMPRAS) || createComprasSheet(ss)
  };
}

function createComprasSheet(ss) {
  const sheet = ss.insertSheet(SHEET_NAME_COMPRAS);
  sheet.appendRow(['Data', 'Nota Fiscal', 'Fornecedor', 'Itens', 'Valor Total', 'Responsável']);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
  return sheet;
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
    else if (page === 'purchases') { data = getPurchasesData(); }
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
      case 'registerPurchase': result = handlePurchaseRegistration(data); break;
      case 'reduceStock': result = handleReduceStock(data); break;
      case 'editItemName': result = handleEditItemName(data); break;
      case 'deleteOrder': result = handleDeleteOrder(data); break;
      case 'deletePurchase': result = handleDeletePurchase(data); break;
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
  const values = pedidos.getRange(2, 1, pedidos.getLastRow() - 1, 10).getValues();
  return values.map(([timestamp, pedidoId, solicitante, setor, idFuncional, item, qtd, justificativa, nivelNecessidade, status]) => ({
    timestamp, pedidoId, solicitante, setor, idFuncional, item, qtd, justificativa, nivelNecessidade, status
  })).reverse();
}

function getPurchasesData() {
  const { compras } = getSheets();
  if (compras.getLastRow() < 2) return [];
  const values = compras.getRange(2, 1, compras.getLastRow() - 1, 6).getValues();
  return values.map(([data, notaFiscal, fornecedor, itens, valorTotal, responsavel]) => ({
    data, notaFiscal, fornecedor, itens, valorTotal, responsavel
  })).reverse();
}

function getAnalyticsData() {
    const { estoque, pedidos } = getSheets();
    const estoqueValues = (estoque.getLastRow() > 1) ? estoque.getRange(2, 2, estoque.getLastRow() - 1, 2).getValues() : [];
    const lowStockItems = estoqueValues.filter(row => Number(row[1]) <= LOW_STOCK_THRESHOLD);
    const highStockItems = estoqueValues.filter(row => Number(row[1]) >= HIGH_STOCK_THRESHOLD);
    const allPedidos = (pedidos.getLastRow() > 1) ? pedidos.getRange(2, 1, pedidos.getLastRow() - 1, 10).getValues() : [];

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
    const pendingOrders = allPedidos.filter(row => row[9] && (row[9] === 'Recebido' || row[9] === 'Em separação'));

    const getTopItems = (orderList) => {
        if (!orderList || orderList.length === 0) return [];
        const counts = orderList.reduce((acc, row) => {
            const item = row[5];
            const qty = Number(row[6]);
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
    pedidos.appendRow([
      timestamp, 
      pedidoId, 
      data.solicitante, 
      data.setor, 
      data.idFuncional, 
      item.nome, 
      item.qtd, 
      data.justificativa, 
      data.nivelNecessidade, 
      "Recebido"
    ]);
  });
  enviarEmailNotificacao(data);
  return { result: 'success' };
}


// =======================================================================
// ↓↓↓ FUNÇÃO MODIFICADA ↓↓↓
// =======================================================================
/**
 * Registra uma nova compra E ATUALIZA O ESTOQUE.
 * Chamado pela action: 'registerPurchase'
 */
function handlePurchaseRegistration(data) {
  // 1. Obter AMBAS as planilhas: Compras (para log) e Estoque (para atualizar)
  const { compras, estoque } = getSheets();
  const dataCompra = data.dataCompra ? new Date(data.dataCompra) : new Date();

  // 2. Salvar o registro na planilha "Compras" (como já fazia)
  const itensStr = JSON.stringify(data.itens);
  compras.appendRow([
    dataCompra,
    data.notaFiscal,
    data.fornecedor,
    itensStr,
    data.valorTotal,
    data.responsavel
  ]);

  // 3. LÓGICA ADICIONADA: Atualizar a planilha "Estoque"
  
  // Pega todos os itens do estoque 1 vez para checar mais rápido
  // Coluna B (nome), C (qtd), D (categoria)
  const allStockRange = estoque.getRange(2, 1, estoque.getLastRow() - 1, 4);
  const allStockValues = allStockRange.getValues(); // [code, itemName, quantity, category]
  
  // Cria um "Mapa" para encontrar itens rápido pelo nome
  const stockItemsMap = new Map();
  allStockValues.forEach((row, index) => {
    // row[1] é o nome do item (coluna B)
    stockItemsMap.set(row[1].toString().trim().toLowerCase(), {
      rowValues: row,     // Guarda os valores da linha
      rowIndex: index + 2 // Guarda o número da linha real (+2 por causa do cabeçalho e índice 0)
    });
  });

  // Loop por cada item que foi comprado
  data.itens.forEach(item => {
    const itemName = item.nome.trim();
    const itemNameLower = itemName.toLowerCase();
    const quantityToAdd = parseInt(item.qtd);
    const category = item.categoria;

    if (stockItemsMap.has(itemNameLower)) {
      // --- CASO 1: O ITEM JÁ EXISTE NO ESTOQUE ---
      const found = stockItemsMap.get(itemNameLower);
      const rowIndex = found.rowIndex;
      const currentStock = parseInt(found.rowValues[2]) || 0; // rowValues[2] é a coluna C (quantidade)
      const newStock = currentStock + quantityToAdd;
      
      // Atualiza a quantidade na Célula C (coluna 3)
      estoque.getRange(rowIndex, 3).setValue(newStock); 
      Logger.log(`Estoque de '${itemName}' (existente) atualizado para ${newStock}.`);

    } else {
      // --- CASO 2: O ITEM É NOVO ---
      const newSKU = generateNewSKU(estoque, category);
      estoque.appendRow([newSKU, itemName, quantityToAdd, category]);
      Logger.log(`Novo item '${itemName}' (SKU: ${newSKU}) adicionado com estoque ${quantityToAdd}.`);
    }
  });
  
  return { result: 'success' };
}
// =======================================================================
// ↑↑↑ FUNÇÃO MODIFICADA ↑↑↑
// =======================================================================


function handleUpdateStatus(data) {
  const { pedidos, estoque } = getSheets();
  if (!data.pedidoId || !data.newStatus) {
    throw new Error("Dados para atualizar status incompletos.");
  }

  const allPedidosRange = pedidos.getRange(2, 1, pedidos.getLastRow() - 1, 10);
  const allPedidosValues = allPedidosRange.getValues();
  let updated = false;

  for (let i = 0; i < allPedidosValues.length; i++) {
    const rowData = allPedidosValues[i];
    const currentPedidoId = rowData[1];

    if (currentPedidoId == data.pedidoId) {
      pedidos.getRange(i + 2, 10).setValue(data.newStatus);
      updated = true;

      const oldStatus = rowData[9];
      if (data.newStatus === 'Concluído' && oldStatus !== 'Concluído') {
        const itemName = rowData[5];
        const itemQty = Number(rowData[6]);
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

// Funções auxiliares
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
  const prefixMap = { 
    'Material de Escritório': 'ESC', 
    'Material para Copa': 'COP', 
    'Material de Limpeza': 'LMP' 
  }; 
  const prefix = prefixMap[category] || 'GER'; 
  // Pega códigos da coluna A (índice 1)
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
  const assunto = `Novo Pedido de Material - Setor ${data.setor} - ${data.nivelNecessidade}`; 
  const corpoHtml = `<h2>Novo Pedido de Material Recebido</h2><p><strong>Solicitante:</strong> ${data.solicitante}</p><p><strong>Setor:</strong> ${data.setor}</p><p><strong>ID Funcional:</strong> ${data.idFuncional || 'Não informado'}</p><p><strong>Nível de Necessidade:</strong> <span style="color: ${data.nivelNecessidade === 'Urgente' ? '#d9534f' : data.nivelNecessidade === 'Alta' ? '#f0ad4e' : '#5cb85c'}; font-weight: bold;">${data.nivelNecessidade}</span></p><hr><h4>Itens Solicitados:</h4><ul>${itensHtml}</ul><p><em>Este e-mail foi gerado automaticamente pelo Sistema de Almoxarifado.</em></p>`; 
  MailApp.sendEmail({ to: EMAIL_DESTINO, subject: assunto, htmlBody: corpoHtml }); 
}

// =======================================================================
// Funções que faltavam (Adicionadas na correção anterior)
// =======================================================================

/**
 * Diminui o estoque de um item manualmente.
 * Chamado pela action: 'reduceStock'
 */
function handleReduceStock(data) {
  const { estoque } = getSheets();
  const values = estoque.getRange(2, 2, estoque.getLastRow() - 1, 2).getValues();
  const itemNameLower = data.itemName.trim().toLowerCase();
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][0].toString().trim().toLowerCase() === itemNameLower) {
      const estoqueAtual = parseInt(values[i][1]) || 0;
      const novoEstoque = estoqueAtual - parseInt(data.quantityToReduce);
      
      if (novoEstoque < 0) {
        throw new Error("Não é possível diminuir mais do que o estoque atual.");
      }
      
      estoque.getRange(i + 2, 3).setValue(novoEstoque); // Coluna C (3) é a quantidade
      return { result: 'success', newStock: novoEstoque };
    }
  }
  throw new Error("Item não encontrado.");
}

/**
 * Edita o nome de um item no estoque.
 * Chamado pela action: 'editItemName'
 */
function handleEditItemName(data) {
  const { estoque } = getSheets();
  // Coluna B (2) é o nome do item
  const values = estoque.getRange(2, 2, estoque.getLastRow() - 1, 1).getValues(); 
  
  // 1. Verifica se o NOVO nome já existe
  const newNameLower = data.newName.trim().toLowerCase();
  if (values.some(row => row[0].toString().trim().toLowerCase() === newNameLower)) {
    throw new Error(`O item "${data.newName}" já existe.`);
  }
  
  // 2. Encontra o nome ANTIGO e atualiza
  const oldNameLower = data.oldName.trim().toLowerCase();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0].toString().trim().toLowerCase() === oldNameLower) {
      estoque.getRange(i + 2, 2).setValue(data.newName.trim()); // Atualiza Coluna B
      return { result: 'success' };
    }
  }
  throw new Error("Item original não encontrado para editar.");
}

/**
 * Exclui um pedido (todas as linhas com o mesmo ID).
 * Chamado pela action: 'deleteOrder'
 */
function handleDeleteOrder(data) {
  const { pedidos } = getSheets();
  // Coluna B (index 1) é o pedidoId
  const allPedidosValues = pedidos.getRange(2, 1, pedidos.getLastRow() - 1, 10).getValues();
  
  let rowsToDelete = [];
  for (let i = 0; i < allPedidosValues.length; i++) {
    if (allPedidosValues[i][1] == data.pedidoId) { 
      rowsToDelete.push(i + 2); // +2 para ajustar índice (começa em 1) e cabeçalho
    }
  }
  
  if (rowsToDelete.length === 0) {
    throw new Error("Pedido não encontrado para exclusão.");
  }
  
  // Exclui as linhas de baixo para cima para não bagunçar os índices
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    pedidos.deleteRow(rowsToDelete[i]);
  }
  
  return { result: 'success' };
}

/**
 * Exclui um registro de compra pela Nota Fiscal.
 * Chamado pela action: 'deletePurchase'
 */
function handleDeletePurchase(data) {
  const { compras } = getSheets();
  // Coluna B (index 2) é a Nota Fiscal
  const nfValues = compras.getRange(2, 2, compras.getLastRow() - 1, 1).getValues(); 
  
  for (let i = 0; i < nfValues.length; i++) {
    if (nfValues[i][0].toString().trim() === data.notaFiscal.trim()) {
      compras.deleteRow(i + 2); // +2 para ajustar índice e cabeçalho
      return { result: 'success' };
    }
  }
  
  throw new Error("Nota Fiscal não encontrada para exclusão.");
}

// =======================================================================
// Função Final
// =======================================================================

function createJsonResponse(data) { 
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); 
}
