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
  sheet.appendRow(['Data de Registro', 'Data da Compra (NF)', 'Nota Fiscal', 'Fornecedor', 'Itens', 'Valor Total', 'Responsável']);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
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
      case 'editPurchase': result = handleEditPurchase(data); break;
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
  const values = pedidos.getRange(2, 1, pedidos.getLastRow() - 1, 11).getValues();
  return values.map((row, index) => ({
    rowId: index + 2,
    timestamp: row[0],
    pedidoId: row[1],
    solicitante: row[2],
    setor: row[3],
    idFuncional: row[4],
    item: row[5],
    qtd: row[6],
    justificativa: row[7],
    nivelNecessidade: row[8],
    status: row[9]
  })).reverse();
}

function getPurchasesData() {
  const { compras } = getSheets();
  if (compras.getLastRow() < 2) return [];
  
  // ✅ CORRIGIDO: Agora lê 7 colunas (A até G)
  const values = compras.getRange(2, 1, compras.getLastRow() - 1, 7).getValues();
  
  return values.map(([dataRegistro, dataCompra, notaFiscal, fornecedor, itens, valorTotal, responsavel], index) => ({
    rowIndex: index + 2,
    dataRegistro,      // Coluna A - Data de Registro
    dataCompra,        // Coluna B - Data da Compra (NF)
    notaFiscal,        // Coluna C - Nota Fiscal
    fornecedor,        // Coluna D - Fornecedor
    itens,             // Coluna E - Itens
    valorTotal,        // Coluna F - Valor Total
    responsavel        // Coluna G - Responsável
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

function handlePurchaseRegistration(data) {
  const { compras, estoque } = getSheets();
  const dataRegistro = new Date(); // Data atual do registro
  const dataCompra = data.dataCompra ? new Date(data.dataCompra) : new Date();

  const itensStr = JSON.stringify(data.itens);
  
  // ✅ CORRIGIDO: Agora adiciona 7 colunas (A até G)
  compras.appendRow([
    dataRegistro,       // Coluna A - Data de Registro
    dataCompra,         // Coluna B - Data da Compra (NF)
    data.notaFiscal,    // Coluna C - Nota Fiscal
    data.fornecedor,    // Coluna D - Fornecedor
    itensStr,           // Coluna E - Itens (JSON)
    data.valorTotal,    // Coluna F - Valor Total
    data.responsavel    // Coluna G - Responsável
  ]);

  const allStockRange = estoque.getRange(2, 1, estoque.getLastRow() - 1, 4);
  const allStockValues = allStockRange.getValues();
  
  const stockItemsMap = new Map();
  allStockValues.forEach((row, index) => {
    stockItemsMap.set(row[1].toString().trim().toLowerCase(), {
      rowValues: row,
      rowIndex: index + 2
    });
  });

  data.itens.forEach(item => {
    const itemName = item.nome.trim();
    const itemNameLower = itemName.toLowerCase();
    const quantityToAdd = parseInt(item.qtd);
    const category = item.categoria;

    if (stockItemsMap.has(itemNameLower)) {
      const found = stockItemsMap.get(itemNameLower);
      const rowIndex = found.rowIndex;
      const currentStock = parseInt(found.rowValues[2]) || 0;
      const newStock = currentStock + quantityToAdd;
      
      estoque.getRange(rowIndex, 3).setValue(newStock);
      Logger.log(`Estoque de '${itemName}' atualizado para ${newStock}.`);
    } else {
      const newSKU = generateNewSKU(estoque, category);
      estoque.appendRow([newSKU, itemName, quantityToAdd, category]);
      Logger.log(`Novo item '${itemName}' (SKU: ${newSKU}) adicionado com estoque ${quantityToAdd}.`);
    }
  });
  
  return { result: 'success' };
}

// =======================================================================
// FUNÇÃO EDITAR NOTA FISCAL
// =======================================================================
function handleEditPurchase(data) {
  const { compras, estoque } = getSheets();
  
  // ✅ CORRIGIDO: Buscar em 7 colunas (A até G)
  const nfValues = compras.getRange(2, 1, compras.getLastRow() - 1, 7).getValues();
  let rowIndex = -1;
  let oldItensStr = '';
  
  for (let i = 0; i < nfValues.length; i++) {
    // Nota Fiscal agora está na coluna C (índice 2)
    if (nfValues[i][2].toString().trim() === data.notaFiscal.trim()) {
      rowIndex = i + 2;
      oldItensStr = nfValues[i][4]; // Itens na coluna E (índice 4)
      break;
    }
  }
  
  if (rowIndex === -1) {
    throw new Error("Nota Fiscal não encontrada.");
  }
  
  // 2. Parsear itens antigos e novos
  let oldItens = [];
  try {
    oldItens = JSON.parse(oldItensStr);
  } catch(e) {
    throw new Error("Erro ao ler itens da nota fiscal original.");
  }
  
  const newItens = data.itens;
  
  // 3. Montar mapa do estoque atual
  const allStockValues = estoque.getRange(2, 1, estoque.getLastRow() - 1, 4).getValues();
  const stockItemsMap = new Map();
  
  allStockValues.forEach((row, index) => {
    const itemNameLower = row[1].toString().trim().toLowerCase();
    stockItemsMap.set(itemNameLower, {
      sku: row[0],
      nome: row[1],
      qtd: parseInt(row[2]) || 0,
      categoria: row[3],
      rowIndex: index + 2
    });
  });
  
  // 4. REVERTER os itens antigos (remover do estoque)
  for (const oldItem of oldItens) {
    const itemNameLower = oldItem.nome.trim().toLowerCase();
    
    if (!stockItemsMap.has(itemNameLower)) {
      throw new Error(`Item '${oldItem.nome}' não encontrado no estoque. Não é possível editar esta nota.`);
    }
    
    const stockItem = stockItemsMap.get(itemNameLower);
    const newQty = stockItem.qtd - parseInt(oldItem.qtd);
    
    if (newQty < 0) {
      throw new Error(`Estoque insuficiente de '${oldItem.nome}'. Atual: ${stockItem.qtd}, Necessário remover: ${oldItem.qtd}. Já foram usadas ${Math.abs(newQty)} unidades.`);
    }
    
    estoque.getRange(stockItem.rowIndex, 3).setValue(newQty);
    stockItem.qtd = newQty;
    Logger.log(`Revertido: '${oldItem.nome}' -${oldItem.qtd} = ${newQty}`);
  }
  
  // 5. ADICIONAR os novos itens (adicionar ao estoque)
  for (const newItem of newItens) {
    const itemNameLower = newItem.nome.trim().toLowerCase();
    const quantityToAdd = parseInt(newItem.qtd);
    
    if (stockItemsMap.has(itemNameLower)) {
      const stockItem = stockItemsMap.get(itemNameLower);
      const newQty = stockItem.qtd + quantityToAdd;
      
      estoque.getRange(stockItem.rowIndex, 3).setValue(newQty);
      stockItem.qtd = newQty;
      Logger.log(`Adicionado: '${newItem.nome}' +${quantityToAdd} = ${newQty}`);
    } else {
      const newSKU = generateNewSKU(estoque, newItem.categoria);
      estoque.appendRow([newSKU, newItem.nome, quantityToAdd, newItem.categoria]);
      
      stockItemsMap.set(itemNameLower, {
        sku: newSKU,
        nome: newItem.nome,
        qtd: quantityToAdd,
        categoria: newItem.categoria,
        rowIndex: estoque.getLastRow()
      });
      
      Logger.log(`Novo item criado: '${newItem.nome}' SKU: ${newSKU} Qtd: ${quantityToAdd}`);
    }
  }
  
  // 6. Atualizar a nota fiscal na planilha de Compras
  const dataCompra = data.dataCompra ? new Date(data.dataCompra) : new Date();
  const newItensStr = JSON.stringify(newItens);
  
  // ✅ CORRIGIDO: Atualizar nas colunas corretas
  compras.getRange(rowIndex, 2).setValue(dataCompra);     // Coluna B - Data da Compra (NF)
  compras.getRange(rowIndex, 4).setValue(data.fornecedor); // Coluna D - Fornecedor
  compras.getRange(rowIndex, 5).setValue(newItensStr);     // Coluna E - Itens
  compras.getRange(rowIndex, 6).setValue(data.valorTotal); // Coluna F - Valor Total
  compras.getRange(rowIndex, 7).setValue(data.responsavel);// Coluna G - Responsável
  
  Logger.log(`Nota Fiscal ${data.notaFiscal} atualizada com sucesso.`);
  
  return { result: 'success' };
}

// =======================================================================
// FUNÇÃO DELETAR NOTA FISCAL - CORRIGIDA E MELHORADA
// =======================================================================
function handleDeletePurchase(data) {
  const { compras, estoque } = getSheets();
  
  // ✅ CORRIGIDO: Ler todas as 7 colunas (A até G)
  const allComprasData = compras.getRange(2, 1, compras.getLastRow() - 1, 7).getValues();
  let rowIndex = -1;
  let itensStr = '';
  
  for (let i = 0; i < allComprasData.length; i++) {
    // Nota Fiscal está na coluna C (índice 2)
    if (allComprasData[i][2].toString().trim() === data.notaFiscal.trim()) {
      rowIndex = i + 2;
      itensStr = allComprasData[i][4]; // Itens na coluna E (índice 4)
      break;
    }
  }
  
  if (rowIndex === -1) {
    throw new Error("Nota Fiscal não encontrada para exclusão.");
  }
  
  Logger.log(`Nota Fiscal encontrada na linha ${rowIndex}`);
  Logger.log(`Itens String: ${itensStr}`);
  
  // 2. Parsear os itens da nota fiscal
  let itens = [];
  try {
    if (!itensStr || itensStr.toString().trim() === '') {
      throw new Error("Campo de itens está vazio.");
    }
    itens = JSON.parse(itensStr);
    
    if (!Array.isArray(itens) || itens.length === 0) {
      throw new Error("Nenhum item encontrado na nota fiscal.");
    }
    
    Logger.log(`${itens.length} itens parseados com sucesso`);
  } catch(e) {
    Logger.log(`Erro ao parsear itens: ${e.toString()}`);
    throw new Error(`Erro ao ler itens da nota fiscal: ${e.message}\n\nA nota pode estar corrompida.`);
  }
  
  // 3. Montar mapa do estoque atual
  const allStockValues = estoque.getRange(2, 1, estoque.getLastRow() - 1, 4).getValues();
  const stockItemsMap = new Map();
  
  allStockValues.forEach((row, index) => {
    const itemNameLower = row[1].toString().trim().toLowerCase();
    stockItemsMap.set(itemNameLower, {
      sku: row[0],
      nome: row[1],
      qtd: parseInt(row[2]) || 0,
      categoria: row[3],
      rowIndex: index + 2
    });
  });
  
  // 4. VALIDAR se é possível remover
  const itemsToRemove = [];
  const insufficientItems = [];
  
  for (const item of itens) {
    const itemNameLower = item.nome.trim().toLowerCase();
    
    if (!stockItemsMap.has(itemNameLower)) {
      throw new Error(`Item '${item.nome}' não encontrado no estoque.`);
    }
    
    const stockItem = stockItemsMap.get(itemNameLower);
    const newQty = stockItem.qtd - parseInt(item.qtd);
    
    if (newQty < 0) {
      insufficientItems.push({
        nome: item.nome,
        estoqueAtual: stockItem.qtd,
        qtdNaNota: item.qtd,
        deficit: Math.abs(newQty)
      });
    } else {
      itemsToRemove.push({
        nome: item.nome,
        qtdRemover: item.qtd,
        novoEstoque: newQty,
        rowIndex: stockItem.rowIndex
      });
    }
  }
  
  // 5. Bloquear se houver estoque insuficiente
  if (insufficientItems.length > 0) {
    let errorMessage = '⚠️ IMPOSSÍVEL EXCLUIR NOTA FISCAL\n\n';
    errorMessage += '❌ Itens com estoque INSUFICIENTE:\n\n';
    
    insufficientItems.forEach(item => {
      errorMessage += `📦 ${item.nome}\n`;
      errorMessage += `   • Estoque: ${item.estoqueAtual}\n`;
      errorMessage += `   • Na NF: ${item.qtdNaNota}\n`;
      errorMessage += `   • Faltam: ${item.deficit}\n\n`;
    });
    
    throw new Error(errorMessage);
  }
  
  // 6. Remover do estoque
  itemsToRemove.forEach(item => {
    estoque.getRange(item.rowIndex, 3).setValue(item.novoEstoque);
    Logger.log(`Removido: ${item.nome} -${item.qtdRemover} = ${item.novoEstoque}`);
  });
  
  // 7. Deletar a nota fiscal
  compras.deleteRow(rowIndex);
  Logger.log(`Nota Fiscal ${data.notaFiscal} excluída.`);
  
  return { result: 'success' };
}

// =======================================================================
// FUNÇÃO ATUALIZAR STATUS (individual por linha)
// =======================================================================
function handleUpdateStatus(data) {
  const { pedidos, estoque } = getSheets();
  
  if (!data.rowId || !data.newStatus) {
    throw new Error("Dados para atualizar status incompletos.");
  }

  const rowData = pedidos.getRange(data.rowId, 1, 1, 10).getValues()[0];
  const oldStatus = rowData[9];
  
  pedidos.getRange(data.rowId, 10).setValue(data.newStatus);
  
  if (data.newStatus === 'Concluído' && oldStatus !== 'Concluído') {
    const itemName = rowData[5];
    const itemQty = Number(rowData[6]);
    Logger.log(`Status 'Concluído' para linha ${data.rowId}, item ${itemName}. Baixa de ${itemQty} unidade(s).`);
    darBaixaEstoque(estoque, itemName, itemQty);
  }

  return { result: 'success' };
}

// =======================================================================
// Funções auxiliares
// =======================================================================
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
      
      estoque.getRange(i + 2, 3).setValue(novoEstoque);
      return { result: 'success', newStock: novoEstoque };
    }
  }
  throw new Error("Item não encontrado.");
}

function handleEditItemName(data) {
  const { estoque } = getSheets();
  const values = estoque.getRange(2, 2, estoque.getLastRow() - 1, 1).getValues(); 
  
  const newNameLower = data.newName.trim().toLowerCase();
  if (values.some(row => row[0].toString().trim().toLowerCase() === newNameLower)) {
    throw new Error(`O item "${data.newName}" já existe.`);
  }
  
  const oldNameLower = data.oldName.trim().toLowerCase();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0].toString().trim().toLowerCase() === oldNameLower) {
      estoque.getRange(i + 2, 2).setValue(data.newName.trim());
      return { result: 'success' };
    }
  }
  throw new Error("Item original não encontrado para editar.");
}

function handleDeleteOrder(data) {
  const { pedidos } = getSheets();
  const allPedidosValues = pedidos.getRange(2, 1, pedidos.getLastRow() - 1, 10).getValues();
  
  let rowsToDelete = [];
  for (let i = 0; i < allPedidosValues.length; i++) {
    if (allPedidosValues[i][1] == data.pedidoId) { 
      rowsToDelete.push(i + 2);
    }
  }
  
  if (rowsToDelete.length === 0) {
    throw new Error("Pedido não encontrado para exclusão.");
  }
  
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    pedidos.deleteRow(rowsToDelete[i]);
  }
  
  return { result: 'success' };
}

function createJsonResponse(data) { 
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); 
}
