const SPREADSHEET_ID = '1-bmYoKpTBcmh9u0WwlxWxy3shH937wAiB-u6Wf9GSmw';
const CONSUMABLES_SPREADSHEET_ID = '1anjxp95VtGD-0YBfw6YYgn2JqZezNcwuNJ1KTEw1GJc';

function parseCurrencyValue(value) {
  return parseFloat((value || '').toString().replace(/[^0-9.]/g, '')) || 0;
}

function normalizeHeader(value) {
  return (value || '').toString().trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function findColumnIndex(headers, aliases, fallbackIndex) {
  for (let i = 0; i < headers.length; i++) {
    if (aliases.indexOf(normalizeHeader(headers[i])) > -1) return i;
  }
  return fallbackIndex;
}

function getConsumablesColumnMap(headers) {
  return {
    id: findColumnIndex(headers, ['id', 'itemid', 'sku', 'code'], 0),
    category: findColumnIndex(headers, ['category'], 1),
    subcategory: findColumnIndex(headers, ['subcategory', 'subcat'], 2),
    description: findColumnIndex(headers, ['description', 'desc', 'itemdescription'], 3),
    size: findColumnIndex(headers, ['size'], 4),
    thickness: findColumnIndex(headers, ['thickness', 'thick'], 5),
    colour: findColumnIndex(headers, ['colour', 'color'], 6),
    cost: findColumnIndex(headers, ['cost', 'unitcost', 'unitprice', 'price', 'priceperunit'], 9),
    stock: findColumnIndex(headers, ['instock', 'stock', 'stocklevel', 'availability', 'status'], 11)
  };
}

function parseConsumableStockStatus(rawValue) {
  const raw = (rawValue || '').toString().trim();
  const normalized = raw.toLowerCase();
  const numericStock = parseFloat(raw);
  const hasNumericStock = !isNaN(numericStock);
  const knownStates = ['high', 'medium', 'low', 'ordered', 'special order'];
  const isKnownState = knownStates.indexOf(normalized) > -1;
  const isSpecialOrder = normalized === 'special order';
  const isOrdered = normalized === 'ordered';
  const canOrderByStatus = isSpecialOrder || normalized === 'high' || normalized === 'medium' || normalized === 'low';
  const canOrder = hasNumericStock ? numericStock > 0 : canOrderByStatus;

  return {
    display: raw || 'Unknown',
    normalized: normalized,
    hasNumericStock: hasNumericStock,
    numericStock: hasNumericStock ? numericStock : null,
    isKnownState: isKnownState,
    isSpecialOrder: isSpecialOrder,
    isOrdered: isOrdered,
    canOrder: canOrder
  };
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('MAC499 Component & Materials Hub')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// SECURITY CONTROLLER
function isStaffUser() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const staffSheet = ss.getSheetByName('Staff');
  const staffData = staffSheet.getDataRange().getValues();
  for (let i = 1; i < staffData.length; i++) {
    if (staffData[i][1].toString().toLowerCase() === email.toLowerCase()) {
      return true;
    }
  }
  return false;
}

// IDENTIFY USER & ROLE
function getUserInfo() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  if (isStaffUser()) {
    return { role: 'staff', email: email };
  }
  
  const studentSheet = ss.getSheetByName('Student Group Numbers');
  const studentData = studentSheet.getDataRange().getValues();
  for (let i = 1; i < studentData.length; i++) {
    if (studentData[i][1].toString().toLowerCase() === email.toLowerCase()) {
      return { role: 'student', email: email, group: studentData[i][2] };
    }
  }
  
  return { role: 'unauthorized', email: email };
}

// GET STUDENT DATA (Pulls from both Storefronts)
function getStudentDashboardData(groupNumber) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // A. Pull Drone Components
  const shopSheet = ss.getSheetByName('Online Shop');
  const shopData = shopSheet.getDataRange().getValues();
  let components = [];
  for(let i=1; i<shopData.length; i++) {
    if(!shopData[i][0]) continue;
    let cleanCost = parseFloat(shopData[i][2].toString().replace(/[^0-9.]/g, '')) || 0;
    components.push({
      item: shopData[i][0],
      stock: parseInt(shopData[i][1]) || 0,
      cost: cleanCost,
      url: shopData[i][5] || '',
      desc: shopData[i][6] || '',
      isConsumable: false
    });
  }
  
  // B. Pull Consumables from external Inventory Spreadsheet based on explicit static columns
  let consumables = [];
  try {
    const consSs = SpreadsheetApp.openById(CONSUMABLES_SPREADSHEET_ID);
    const consSheet = consSs.getSheetByName('Inventory');
    if (consSheet) {
      const consData = consSheet.getDataRange().getValues();
      const headers = consData.length > 0 ? consData[0] : [];
      const colMap = getConsumablesColumnMap(headers);
      for (let i = 1; i < consData.length; i++) {
        if (!consData[i][colMap.id]) continue; // Skip if ID is empty
        
        let id = consData[i][colMap.id].toString();
        let category = (consData[i][colMap.category] || '').toString();
        let subcategory = (consData[i][colMap.subcategory] || '').toString();
        let description = (consData[i][colMap.description] || '').toString();
        
        // Extract basic dimensions
        let size = consData[i][colMap.size] ? consData[i][colMap.size].toString().trim() : '';
        let thickness = consData[i][colMap.thickness] ? consData[i][colMap.thickness].toString().trim() : '';
        let colour = consData[i][colMap.colour] ? consData[i][colMap.colour].toString().trim() : '';
        
        let dimsArray = [];
        if (size) dimsArray.push("Size: " + size);
        if (thickness) dimsArray.push("Thick: " + thickness);
        if (colour) dimsArray.push("Colour: " + colour);
        let dimensionsString = dimsArray.length > 0 ? dimsArray.join(' | ') : '—';
        let consumableLabel = (subcategory || 'General') + " | " + dimensionsString;
        
        let cleanCost = parseCurrencyValue(consData[i][colMap.cost]);
        let stockMeta = parseConsumableStockStatus(consData[i][colMap.stock]);

        consumables.push({
          item: id,
          category: category,
          subcategory: subcategory,
          desc: description,
          dimensions: dimensionsString,
          stock: stockMeta.display,
          stockStatus: stockMeta.normalized,
          numericStock: stockMeta.numericStock,
          canOrder: stockMeta.canOrder,
          isSpecialOrder: stockMeta.isSpecialOrder,
          label: consumableLabel,
          cost: cleanCost,
          isConsumable: true
        });
      }
    }
  } catch (err) {
    Logger.log("Consumables fetch exception: " + err.message);
  }

  // C. Calculate Budgets & Read current BoM
  const budgetSheet = ss.getSheetByName('Budgets');
  const budgetData = budgetSheet.getDataRange().getValues();
  let totalBudget = 0;
  for(let i=1; i<budgetData.length; i++) {
    if(budgetData[i][0].toString() === groupNumber.toString()) {
      totalBudget = parseFloat(budgetData[i][1].toString().replace(/[^0-9.]/g, '')) || 0;
    }
  }
  
  const bomSheet = ss.getSheetByName('Group ' + groupNumber + ' BoM');
  let spent = 0;
  let bomItems = [];
  if(bomSheet) {
    const bomData = bomSheet.getDataRange().getValues();
    for(let i=1; i<bomData.length; i++) {
      if(!bomData[i][1]) continue;
      let costVal = parseFloat(bomData[i][3].toString().replace(/[^0-9.]/g, '')) || 0;
      spent += costVal;
      bomItems.push({ 
        date: bomData[i][0] instanceof Date ? bomData[i][0].toLocaleDateString() : bomData[i][0].toString(), 
        item: bomData[i][1], 
        qty: bomData[i][2], 
        cost: costVal 
      });
    }
  }
  
  return { components: components, consumables: consumables, budget: totalBudget, spent: spent, remaining: totalBudget - spent, bom: bomItems };
}

// TRANSACTION PLACE ORDER (Handles split stock updates & filtered email logs)
function placeOrder(orderItems, groupNumber) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000); 
  
  try {
    const studentEmail = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const consSs = SpreadsheetApp.openById(CONSUMABLES_SPREADSHEET_ID);
    
    const budgetSheet = ss.getSheetByName('Budgets');
    const budgetData = budgetSheet.getDataRange().getValues();
    let totalBudget = 0;
    for(let i=1; i<budgetData.length; i++) {
      if(budgetData[i][0].toString() === groupNumber.toString()) {
        totalBudget = parseFloat(budgetData[i][1].toString().replace(/[^0-9.]/g, '')) || 0;
      }
    }
    
    const bomSheet = ss.getSheetByName('Group ' + groupNumber + ' BoM');
    let spent = 0;
    if(bomSheet) {
      const bomData = bomSheet.getDataRange().getValues();
      for(let i=1; i<bomData.length; i++) {
        if(!bomData[i][1]) continue;
        spent += parseFloat(bomData[i][3].toString().replace(/[^0-9.]/g, '')) || 0;
      }
    }
    let remainingBudget = totalBudget - spent;
    
    const shopSheet = ss.getSheetByName('Online Shop');
    const shopData = shopSheet.getDataRange().getValues();
    
    const consSheet = consSs.getSheetByName('Inventory');
    const consData = consSheet ? consSheet.getDataRange().getValues() : [];
    const consHeaders = consData.length > 0 ? consData[0] : [];
    const consColMap = getConsumablesColumnMap(consHeaders);
    
    const orderSheet = ss.getSheetByName('Order History');
    
    let totalOrderCost = 0;
    let componentDetailsText = "";
    let hasComponents = false;
    let inventoryUpdates = [];
    
    for(let k=0; k<orderItems.length; k++) {
      let clientItem = orderItems[k];
      let found = false;
      
      if (clientItem.isConsumable) {
        if (consData.length <= 1) throw new Error("Consumables Inventory database sheet offline.");

        for (let i = 1; i < consData.length; i++) {
          if (consData[i][consColMap.id].toString() === clientItem.item) { // Match by ID
            found = true;
            const stockMeta = parseConsumableStockStatus(consData[i][consColMap.stock]);
            if (!stockMeta.canOrder) throw new Error("Consumable currently unavailable for ordering: " + clientItem.item);
            if (stockMeta.hasNumericStock && stockMeta.numericStock < clientItem.qty) {
              throw new Error("Insufficient stock for consumable material: " + clientItem.item);
            }

            let subcategory = (consData[i][consColMap.subcategory] || '').toString();
            let size = consData[i][consColMap.size] ? consData[i][consColMap.size].toString().trim() : '';
            let thickness = consData[i][consColMap.thickness] ? consData[i][consColMap.thickness].toString().trim() : '';
            let colour = consData[i][consColMap.colour] ? consData[i][consColMap.colour].toString().trim() : '';
            let dimsArray = [];
            if (size) dimsArray.push("Size: " + size);
            if (thickness) dimsArray.push("Thick: " + thickness);
            if (colour) dimsArray.push("Colour: " + colour);
            let dimensionsString = dimsArray.length > 0 ? dimsArray.join(' | ') : '—';
            let descriptionLabel = (subcategory || 'General') + " | " + dimensionsString;

            let cleanCost = parseCurrencyValue(consData[i][consColMap.cost]);
            let itemTotalCost = cleanCost * clientItem.qty;
            totalOrderCost += itemTotalCost;

            inventoryUpdates.push({
              targetSheet: stockMeta.hasNumericStock ? consSheet : null,
              row: stockMeta.hasNumericStock ? i + 1 : null,
              col: stockMeta.hasNumericStock ? (consColMap.stock + 1) : null,
              newStock: stockMeta.hasNumericStock ? (stockMeta.numericStock - clientItem.qty) : null,
              item: clientItem.item,
              qty: clientItem.qty,
              cost: itemTotalCost,
              isConsumable: true,
              isSpecialOrderConsumable: stockMeta.isSpecialOrder,
              descriptionLabel: descriptionLabel
            });
            break;
          }
        }
      } else {
        for(let i=1; i<shopData.length; i++) {
          if(shopData[i][0] === clientItem.item) {
            found = true;
            let currentStock = parseInt(shopData[i][1]);
            if(currentStock < clientItem.qty) throw new Error("Insufficient stock for component: " + clientItem.item);
            
            let cleanCost = parseFloat(shopData[i][2].toString().replace(/[^0-9.]/g, '')) || 0;
            let itemTotalCost = cleanCost * clientItem.qty;
            totalOrderCost += itemTotalCost;
            
            inventoryUpdates.push({
              targetSheet: shopSheet, row: i+1, col: 2, newStock: currentStock - clientItem.qty,
              item: clientItem.item, qty: clientItem.qty, cost: itemTotalCost, isConsumable: false
            });
            break;
          }
        }
      }
      if(!found) throw new Error("Item profile records missing: " + clientItem.item);
    }
    
    if(totalOrderCost > remainingBudget) throw new Error("Order total financial impact exceeds group remaining budget allocation.");
    
    // Commit adjustments & register timeline entries
    let hasSpecialOrderConsumables = false;
    let specialOrderDetailsText = "";

    inventoryUpdates.forEach(u => {
      if (!u.isConsumable && u.targetSheet && u.row && u.col) {
        u.targetSheet.getRange(u.row, u.col).setValue(u.newStock);
      }

      let orderId = "ORD-" + new Date().getTime() + "-" + Math.floor(Math.random()*1000);
      
      let finalLabel = u.isConsumable ? "[Consumable] ID: " + u.item + " | " + (u.descriptionLabel || 'General | —') : u.item;
      let initialStatus = u.isConsumable ? (u.isSpecialOrderConsumable ? "Pending" : "Fulfilled") : "Pending";
      
      orderSheet.appendRow([orderId, new Date(), groupNumber, studentEmail, finalLabel, u.qty, u.cost, initialStatus]);
      if(bomSheet) bomSheet.appendRow([new Date(), finalLabel, u.qty, u.cost]);
      
      if(!u.isConsumable) {
        hasComponents = true;
        componentDetailsText += `${u.qty}x ${u.item} (£${u.cost.toFixed(2)})\n`;
      } else if (u.isSpecialOrderConsumable) {
        hasSpecialOrderConsumables = true;
        specialOrderDetailsText += `${u.qty}x ID: ${u.item} | ${u.descriptionLabel || 'General | —'} (£${u.cost.toFixed(2)})\n`;
      }
    });
    
    // Alert staff when fulfillment action is required.
    if(hasComponents || hasSpecialOrderConsumables) {
      const staffEmails = ss.getSheetByName('Staff').getDataRange().getValues().slice(1).map(row => row[1]).join(',');
      if(staffEmails) {
        let emailBody = `Group ${groupNumber} (${studentEmail}) has submitted a new order.\n\n`;
        if (hasComponents) {
          emailBody += `Drone Components Requiring Fulfillment:\n${componentDetailsText}\n`;
        }
        if (hasSpecialOrderConsumables) {
          emailBody += `Special Order Consumables Requiring Request/Approval:\n${specialOrderDetailsText}\n`;
        }
        emailBody += `Total Ordered: £${totalOrderCost.toFixed(2)}`;

        MailApp.sendEmail({
          to: staffEmails,
          subject: `MAC499 New Order Requiring Action - Group ${groupNumber}`,
          body: emailBody
        });
      }
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// EXTENDED STAFF DATA RETRIEVAL
function getStaffDashboardData() {
  if (!isStaffUser()) throw new Error("Access Denied");
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const orderSheet = ss.getSheetByName('Order History');
  const orderData = orderSheet.getDataRange().getValues();
  let orders = [];
  for(let i=1; i<orderData.length; i++) {
    if(!orderData[i][0]) continue;
    orders.push({
      row: i+1, id: orderData[i][0],
      date: orderData[i][1] instanceof Date ? orderData[i][1].toLocaleString() : orderData[i][1].toString(),
      group: orderData[i][2], email: orderData[i][3] || 'System',
      item: orderData[i][4], qty: orderData[i][5],
      total: parseFloat(orderData[i][6] || 0).toFixed(2), status: orderData[i][7]
    });
  }
  orders.reverse();
  
  const budgetSheet = ss.getSheetByName('Budgets');
  const budgetData = budgetSheet.getDataRange().getValues();
  let budgets = [];
  for(let i=1; i<budgetData.length; i++) {
    if(!budgetData[i][0]) continue;
    budgets.push({ group: budgetData[i][0], budget: budgetData[i][1] });
  }
  
  const rosterSheet = ss.getSheetByName('Student Group Numbers');
  const rosterData = rosterSheet.getDataRange().getValues();
  let roster = [];
  for(let i=1; i<rosterData.length; i++) {
    if(!rosterData[i][1]) continue;
    roster.push({ row: i+1, name: rosterData[i][0], email: rosterData[i][1], group: rosterData[i][2] });
  }
  
  let boms = {};
  const sheets = ss.getSheets();
  sheets.forEach(sh => {
    let name = sh.getName();
    if(name.indexOf('Group ') === 0 && name.indexOf(' BoM') > -1) {
      let gNum = name.replace('Group ', '').replace(' BoM', '').trim();
      let data = sh.getDataRange().getValues();
      let items = [];
      for(let i=1; i<data.length; i++) {
        if(!data[i][1]) continue;
        items.push({
          date: data[i][0] instanceof Date ? data[i][0].toLocaleDateString() : data[i][0].toString(),
          item: data[i][1], qty: data[i][2], cost: parseFloat(data[i][3] || 0).toFixed(2)
        });
      }
      boms[gNum] = items;
    }
  });

  const shopSheet = ss.getSheetByName('Online Shop');
  const shopData = shopSheet.getDataRange().getValues();
  let shopItems = [];
  for(let i=1; i<shopData.length; i++) {
    if(!shopData[i][0]) continue;
    shopItems.push({
      row: i+1, item: shopData[i][0], stock: parseInt(shopData[i][1]) || 0,
      cost: parseFloat(shopData[i][2].toString().replace(/[^0-9.]/g, '')) || 0,
      category: shopData[i][3] || '', subcategory: shopData[i][4] || '',
      url: shopData[i][5] || '', desc: shopData[i][6] || ''
    });
  }
  
  return { orders: orders, budgets: budgets, roster: roster, boms: boms, shopItems: shopItems };
}

function markOrderFulfilled(rowNumber) {
  if (!isStaffUser()) throw new Error("Access Denied");
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Order History').getRange(rowNumber, 8).setValue('Fulfilled');
  return true;
}

function updateBudget(groupNumber, newBudget) {
  if (!isStaffUser()) throw new Error("Access Denied");
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Budgets');
  const data = sheet.getDataRange().getValues();
  let numericBudget = parseFloat(newBudget.toString().replace(/[^0-9.]/g, ''));
  
  for(let i=1; i<data.length; i++) {
    if(data[i][0].toString() === groupNumber.toString()) {
      sheet.getRange(i+1, 2).setValue(numericBudget);
      return { success: true };
    }
  }
  sheet.appendRow([groupNumber, numericBudget]);
  return { success: true };
}

function saveComponent(isEdit, row, item, stock, cost, category, subcategory, url, desc) {
  if (!isStaffUser()) throw new Error("Access Denied");
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Online Shop');
  let numericCost = parseFloat(cost.toString().replace(/[^0-9.]/g, '')) || 0;
  
  if (isEdit && row) {
    sheet.getRange(row, 1, 1, 7).setValues([[item, parseInt(stock), numericCost, category, subcategory, url, desc]]);
    return { success: true, action: 'updated' };
  } else {
    sheet.appendRow([item, parseInt(stock), numericCost, category, subcategory, url, desc]);
    return { success: true, action: 'added' };
  }
}

function upsertStudentRoster(name, email, group) {
  if (!isStaffUser()) throw new Error("Access Denied");
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Student Group Numbers');
  const data = sheet.getDataRange().getValues();
  const cleanEmail = email.toString().trim().toLowerCase();
  
  for(let i=1; i<data.length; i++) {
    if(data[i][1].toString().trim().toLowerCase() === cleanEmail) {
      sheet.getRange(i+1, 1).setValue(name);
      sheet.getRange(i+1, 3).setValue(group);
      return { success: true, action: 'updated' };
    }
  }
  sheet.appendRow([name, email.toString().trim(), group]);
  return { success: true, action: 'added' };
}

function createNewGroup(groupNumber, initialBudget) {
  if (!isStaffUser()) throw new Error("Access Denied");
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const budgetSheet = ss.getSheetByName('Budgets');
  const budgetData = budgetSheet.getDataRange().getValues();
  const targetGroupStr = groupNumber.toString().trim();
  let numericBudget = parseFloat(initialBudget.toString().replace(/[^0-9.]/g, '')) || 0;
  
  for(let i=1; i<budgetData.length; i++) {
    if(budgetData[i][0].toString().trim() === targetGroupStr) {
      return { success: false, error: "Group ID code matches existing entity tracking logs." };
    }
  }
  
  budgetSheet.appendRow([targetGroupStr, numericBudget]);
  const newSheetName = 'Group ' + targetGroupStr + ' BoM';
  let targetSheet = ss.getSheetByName(newSheetName);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(newSheetName);
    targetSheet.appendRow(['Date', 'Item', 'Quantity', 'Cost']);
    targetSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e0e0e0');
  }
  return { success: true };
}