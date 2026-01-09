/**
 * Web 应用入口
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('关系链管理系统')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 获取页面内容（动态加载页面）
 * @param {string} pageId - 页面ID (tables, relationship, statistics, settings)
 */
function getPageContent(pageId) {
  try {
    let fileName = '';
    switch(pageId) {
      case 'tables':
        fileName = 'pages-TablesPage';
        break;
      case 'relationship':
        fileName = 'pages-RelationshipPage';
        break;
      case 'statistics':
        fileName = 'pages-StatisticsPage';
        break;
      case 'settings':
        fileName = 'pages-SettingsPage';
        break;
      default:
        return '<div style="padding:20px; color:#dc2626;">页面不存在</div>';
    }
    
    const html = HtmlService.createTemplateFromFile(fileName).getRawContent();
    return html;
  } catch (e) {
    Logger.log('getPageContent 错误: ' + e.toString());
    return '<div style="padding:20px; color:#dc2626;">页面加载失败: ' + e.toString() + '</div>';
  }
}

/**
 * 获取 VT 工作表的客户数据
 * 修复：现在会返回 object 数组并包含 originalIndex 以供删除/更新定位
 */
function getCustomerData() {
  try {
    const SPREADSHEET_ID = '1ibTwstvYB2x2_uLL3_wH6sdBvaH5ixTDzPLJKmxAmv0';
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('VT');
    if (!sheet) {
      const msg = '找不到 VT 工作表';
      Logger.log(msg);
      return { error: msg };
    }
    const data = sheet.getDataRange().getValues();
    Logger.log('getCustomerData 行数: ' + data.length);

    const headers = (data.length > 0) ? data[0].map(h => h === null ? '' : String(h)) : [];

    // 修复：返回结构化的对象数组，确保前端 TablePage 的 renderTableContent 能正确获取索引
    const rows = data.slice(1).map((row, index) => {
      return {
        originalIndex: index, // 关键：保存这一行在原始数组中的位置 (0对应第2行)
        date: row[0] ? (Object.prototype.toString.call(row[0]) === '[object Date]' ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'M月d日') : String(row[0])) : "",
        name: String(row[1] || ""),
        age: String(row[2] || ""),
        phone: String(row[3] || ""),
        address: String(row[4] || ""),
        recommender: String(row[5] || ""),
        remarks: String(row[6] || ""),
        bankCard: String(row[7] || ""),
        verifiedCard: String(row[8] || ""),
        usedCard: String(row[9] || "")
      };
    });

    return rows; // 直接返回数组，适配 TablesPage 的 SuccessHandler(data)
  } catch (e) {
    Logger.log('getCustomerData 错误: ' + e.toString());
    return { error: e.toString() };
  }
}

/**
 * 获取关系链视图数据
 */
function getRelationshipChainData() {
  try {
    const SPREADSHEET_ID = '1ibTwstvYB2x2_uLL3_wH6sdBvaH5ixTDzPLJKmxAmv0';
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    const allSheets = spreadsheet.getSheets();
    const sheetNames = allSheets.map(s => s.getName());
    Logger.log('可用的工作表: ' + sheetNames.join(', '));
    
    let sheet = spreadsheet.getSheetByName('VT');
    if (!sheet) {
      Logger.log('未找到 VT 工作表，尝试查找其他名称');
      sheet = allSheets.find(s => s.getName().includes('VT'));
      if (!sheet) {
        const msg = '找不到 VT 工作表。可用工作表：' + sheetNames.join(', ');
        Logger.log(msg);
        return { error: msg };
      }
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log('获取的数据行数: ' + data.length);
    
    if (data.length === 0) {
      const msg = 'VT 工作表为空';
      Logger.log(msg);
      return { error: msg };
    }
    
    const headers = (data.length > 0) ? data[0].map(h => h === null ? '' : String(h)) : [];
    
    const rows = data.slice(1).map(row => row.map(cell => {
      if (cell === null || typeof cell === 'undefined') return '';
      if (Object.prototype.toString.call(cell) === '[object Date]') {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      return String(cell);
    }));
    
    const customers = [];
    const customerMap = {}; 
    
    rows.forEach((row, index) => {
      const customer = {
        id: index,
        date: row[0],      // A
        name: row[1],      // B
        age: row[2],       // C
        phone: row[3],     // D
        address: row[4],   // E
        recommender: row[5], // F
        remarks: row[6],   // G
        bankCard: row[7],  // H
        verifiedCard: row[8], // I
        usedCard: row[9],   // J
        originalIndex: index // 重要：保留原始索引
      };
      customers.push(customer);
      
      const recommender = customer.recommender || '无推荐人';
      if (!customerMap[recommender]) {
        customerMap[recommender] = [];
      }
      customerMap[recommender].push(customer);
    });
    
    return {
      customers: customers,
      customerMap: customerMap,
      headers: headers
    };
  } catch (e) {
    Logger.log('getRelationshipChainData 错误: ' + e.toString());
    return { error: e.toString() };
  }
}

/**
 * 获取单个客户的详细信息和关系链
 */
function getCustomerDetail(customerId) {
  const relationshipData = getRelationshipChainData();
  if (relationshipData && relationshipData.error) {
    return { error: relationshipData.error };
  }
  const customers = relationshipData.customers;

  if (customerId < 0 || customerId >= customers.length) {
    return { error: '客户不存在' };
  }

  const customer = customers[customerId];
  
  let superior = null;
  if (customer.recommender) {
    const superiorList = customers.filter(c => c.name === customer.recommender);
    if (superiorList.length > 0) {
      superior = superiorList[0];
    }
  }
  
  const subordinates = customers.filter(c => c.recommender === customer.name);
  
  return {
    customer: customer,
    superior: superior,
    subordinates: subordinates
  };
}

/**
 * 调试用
 */
function pingSpreadsheet() {
  const SPREADSHEET_ID = '1ibTwstvYB2x2_uLL3_wH6sdBvaH5ixTDzPLJKmxAmv0';
  try {
    const eff = Session.getEffectiveUser && Session.getEffectiveUser().getEmail ? Session.getEffectiveUser().getEmail() : 'unknown';
    const active = Session.getActiveUser && Session.getActiveUser().getEmail ? Session.getActiveUser().getEmail() : 'unknown';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetNames = ss.getSheets().map(s => s.getName());
    return { ok: true, effectiveUser: eff, activeUser: active, sheetNames: sheetNames };
  } catch (e) {
    const eff = (Session.getEffectiveUser && Session.getEffectiveUser().getEmail) ? Session.getEffectiveUser().getEmail() : 'unknown';
    return { ok: false, error: e.toString(), effectiveUser: eff };
  }
}

/**
 * 更新单个客户数据
 */
function updateCustomerData(rowIndex, updatedCustomer) {
  try {
    const SPREADSHEET_ID = '1ibTwstvYB2x2_uLL3_wH6sdBvaH5ixTDzPLJKmxAmv0';
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('VT');
    
    if (!sheet) {
      return { ok: false, error: '找不到 VT 工作表' };
    }
    
    const actualRow = rowIndex + 2;
    const lastRow = sheet.getLastRow();
    
    if (actualRow < 2 || actualRow > lastRow) {
      return { ok: false, error: '行号超出范围' };
    }
    
    const vals = [[
      updatedCustomer.date || '',       
      updatedCustomer.name || '',       
      updatedCustomer.age || '',        
      updatedCustomer.phone || '',      
      updatedCustomer.address || '',    
      updatedCustomer.recommender || '', 
      updatedCustomer.remarks || '',    
      updatedCustomer.bankCard || '',   
      updatedCustomer.verifiedCard || '', 
      updatedCustomer.usedCard || ''    
    ]];
    
    sheet.getRange(actualRow, 1, 1, 10).setValues(vals);
    return { ok: true, message: '数据已保存' };
  } catch (e) {
    Logger.log('updateCustomerData 错误: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}

/**
 * 添加新客户
 */
function addNewCustomer(customer) {
  try {
    const SPREADSHEET_ID = '1ibTwstvYB2x2_uLL3_wH6sdBvaH5ixTDzPLJKmxAmv0';
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('VT');
    
    if (!sheet) {
      return { ok: false, error: '找不到 VT 工作表' };
    }
    
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    
    const vals = [[
      customer.date || '',
      customer.name || '',
      customer.age || '',
      customer.phone || '',
      customer.address || '',
      customer.recommender || '',
      customer.remarks || '',
      customer.bankCard || '',
      customer.verifiedCard || '',
      customer.usedCard || ''
    ]];
    
    sheet.getRange(newRow, 1, 1, 10).setValues(vals);
    return { ok: true, message: '客户已添加' };
  } catch (e) {
    Logger.log('addNewCustomer 错误: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}

/**
 * 核心新增：删除客户数据函数
 */
function deleteCustomerRow(rowIndex) {
  try {
    const SPREADSHEET_ID = '1ibTwstvYB2x2_uLL3_wH6sdBvaH5ixTDzPLJKmxAmv0';
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('VT');
    if (!sheet) return { ok: false, error: '找不到 VT 工作表' };
    
    const actualRow = rowIndex + 2;
    sheet.deleteRow(actualRow);
    return { ok: true };
  } catch (e) {
    Logger.log('deleteCustomerRow 错误: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}

/**
 * 从文本解析客户信息并添加到表格
 */
function fillTable(text) {
  try {
    if (!text) return { ok: false, error: '输入文本不能为空' };
    
    const lines = text.trim().split(/\r?\n/);
    const data = {};
    
    lines.forEach(line => {
      const colonIndex = line.indexOf('：') !== -1 ? line.indexOf('：') : line.indexOf(':');
      if (colonIndex !== -1) {
        const key = line.substring(0, colonIndex).trim();
        const value = line.substring(colonIndex + 1).trim();
        
        switch(key) {
          case '姓名': data.name = value; break;
          case '年龄': data.age = value; break;
          case '手机号': data.phone = value; break;
          case '推荐人': data.recommender = value; break;
          case '目前居住地':
          case '家庭住址': data.address = value; break;
          case '家庭情况简单描述': data.remarks = value; break;
        }
      }
    });
    
    if (!data.name || !data.phone) {
      return { ok: false, error: '姓名和手机号是必填项' };
    }
    
    const SPREADSHEET_ID = '1ibTwstvYB2x2_uLL3_wH6sdBvaH5ixTDzPLJKmxAmv0';
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('VT');
    
    if (!sheet) {
      return { ok: false, error: '找不到 VT 工作表' };
    }
    
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    
    const today = new Date();
    const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'M月d日');
    
    const finalRecommender = (data.recommender && data.recommender.trim() !== "") ? data.recommender : "电销";
    data.recommender = finalRecommender;
    data.date = dateStr;

    const vals = [[
      dateStr,
      data.name || '',
      data.age || '',
      data.phone || '',
      data.address || '',
      finalRecommender,
      data.remarks || '',
      '', '', ''
    ]];
    
    sheet.getRange(newRow, 1, 1, 10).setValues(vals);
    return { ok: true, row: newRow, data: data, message: '客户信息已填充' };
  } catch (e) {
    Logger.log('fillTable 错误: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}

/**
 * 获取当前用户的邮箱地址
 */
function getCurrentUserEmail() {
  try {
    const user = Session.getEffectiveUser();
    const email = user ? user.getEmail() : 'unknown';
    return { ok: true, email: email };
  } catch (e) {
    Logger.log('getCurrentUserEmail 错误: ' + e.toString());
    return { ok: false, email: 'unknown' };
  }
}
