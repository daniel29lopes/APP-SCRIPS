const getActiveSpreadsheet = () => {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    // Return null if running outside GAS environment (for local syntax checks)
    return null;
  }
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('TRACK-FAB ERP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitialData() {
  try {
    const ss = getActiveSpreadsheet();
    if (!ss) return { error: "Sem acesso ao SpreadsheetApp" };

    const getSheetData = (sheetName) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return [];
      // Using getDisplayValues() instead of getValues() to prevent cell error / date serialization issues
      const data = sheet.getDataRange().getDisplayValues();
      if (data.length <= 1) return []; // Empty or only headers

      const headers = data[0];
      const sanitizedData = [];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        // Skip ghost rows (completely empty lines)
        if (row.join('').trim() === '') continue;

        let obj = {};
        headers.forEach((header, index) => {
          obj[header] = row[index];
        });
        obj['_rowIndex'] = i + 1; // +1 for 0-index
        sanitizedData.push(obj);
      }

      return sanitizedData;
    };

    return {
      db1: getSheetData('DB_1_SpoolTracker_Definitivo'),
      db2: getSheetData('DB_2_WeldingMap_NDT'),
      db3: getSheetData('DB_3_BOM_Engenharia'),
      db4: getSheetData('DB_4_Inventario_Logistica'),
      db5: getSheetData('DB_5_Planeamento_Faturacao'),
      db6: getSheetData('DB_6_Logs')
    };
  } catch (e) {
    return { error: e.message };
  }
}

function logAction(action, target, details) {
  const ss = getActiveSpreadsheet();
  if (!ss) return;
  let sheet = ss.getSheetByName('DB_6_Logs');
  if (!sheet) {
    sheet = ss.insertSheet('DB_6_Logs');
    sheet.appendRow(['Data_Hora', 'Utilizador', 'Acao_Realizada', 'Alvo', 'Detalhes']);
  }

  let user = "Utilizador Desconhecido";
  try {
    user = Session.getEffectiveUser().getEmail() || "Utilizador Ativo";
  } catch(e) {
    user = "Utilizador Ativo";
  }

  const now = new Date();
  sheet.appendRow([now, user, action, target, details]);
}

function deallocateMaterial(tagSpool, itemNo, qty) {
  const ss = getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_3_BOM_Engenharia');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const hSpool = headers.indexOf('Tag_Spool');
  const hItem = headers.indexOf('Item_No');
  const hQty = headers.indexOf('Qtd_Alocada');

  for (let i = 1; i < data.length; i++) {
    if (data[i][hSpool] == tagSpool && data[i][hItem] == itemNo) {
      const currentQty = Number(data[i][hQty]);
      if (currentQty > qty) {
        sheet.getRange(i + 1, hQty + 1).setValue(currentQty - qty);
      } else {
        sheet.deleteRow(i + 1);
      }
      logAction("Desalocação Material", tagSpool, `Item ${itemNo} devolvido (Qtd: ${qty})`);
      break;
    }
  }
  return getInitialData();
}

function allocateMaterial(tagIso, itemNo, tagSpool, qty, destination) {
  const ss = getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_3_BOM_Engenharia');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const hIso = headers.indexOf('Tag_ISO');
  const hItem = headers.indexOf('Item_No');
  const hSpool = headers.indexOf('Tag_Spool');

  // Find original line (where Tag_Spool is empty)
  let originalRow = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][hIso] == tagIso && data[i][hItem] == itemNo && (!data[i][hSpool] || data[i][hSpool] === '')) {
      originalRow = data[i];
      break;
    }
  }

  if (!originalRow) {
      throw new Error("Linha base do material não encontrada. Pode já ter sido totalmente alocada de outra forma, ou os dados originais não estão corretos.");
  }

  let newRow = [...originalRow];
  newRow[hSpool] = tagSpool;
  newRow[headers.indexOf('Qtd_Alocada')] = qty;
  newRow[headers.indexOf('Destino')] = destination;

  sheet.appendRow(newRow);

  logAction("Alocação Material", tagSpool, `Item ${itemNo} alocado ${qty} para ${destination} (ISO: ${tagIso})`);
  return getInitialData();
}

function createJoint(jointData) {
  const ss = getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_2_WeldingMap_NDT');
  const headers = sheet.getDataRange().getValues()[0];

  let newRow = new Array(headers.length).fill('');
  Object.keys(jointData).forEach(key => {
    let index = headers.indexOf(key);
    if (index !== -1) {
      newRow[index] = jointData[key];
    }
  });

  sheet.appendRow(newRow);
  logAction("Criada Junta", jointData.Tag_Spool, `Junta ${jointData.ID_Junta} criada com sucesso.`);
  return getInitialData();
}

function deleteSpool(tagSpool) {
  const ss = getActiveSpreadsheet();
  const sheetSpool = ss.getSheetByName('DB_1_SpoolTracker_Definitivo');
  const spoolData = sheetSpool.getDataRange().getValues();
  const hSpool = spoolData[0].indexOf('Tag_Spool');

  for (let i = 1; i < spoolData.length; i++) {
    if (spoolData[i][hSpool] == tagSpool) {
      sheetSpool.deleteRow(i + 1);
      logAction("Spool Apagado", tagSpool, `Spool apagado do sistema.`);
      break;
    }
  }

  // Deallocate all associated materials
  const sheetBOM = ss.getSheetByName('DB_3_BOM_Engenharia');
  const bomData = sheetBOM.getDataRange().getValues();
  const hSpoolBOM = bomData[0].indexOf('Tag_Spool');

  for (let i = bomData.length - 1; i >= 1; i--) {
    if (bomData[i][hSpoolBOM] == tagSpool) {
      sheetBOM.deleteRow(i + 1);
    }
  }

  return getInitialData();
}

function updateSpoolsStatus(spoolTags, newStatus) {
  const ss = getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_1_SpoolTracker_Definitivo');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const hTag = headers.indexOf('Tag_Spool');
  const hStatus = headers.indexOf('Estado_Fabrico');

  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    if (spoolTags.includes(data[i][hTag])) {
      sheet.getRange(i + 1, hStatus + 1).setValue(newStatus);
      updated++;
    }
  }

  logAction("Mudança de Estado", "Múltiplos Spools", `Estado alterado para ${newStatus} em ${updated} spools.`);
  return getInitialData();
}

function updateSpoolDetails(tagSpool, newStatus, newNotes) {
  const ss = getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_1_SpoolTracker_Definitivo');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const hTag = headers.indexOf('Tag_Spool');
  const hStatus = headers.indexOf('Estado_Fabrico');
  const hNotes = headers.indexOf('Observacoes');

  for (let i = 1; i < data.length; i++) {
    if (data[i][hTag] == tagSpool) {
      sheet.getRange(i + 1, hStatus + 1).setValue(newStatus);
      sheet.getRange(i + 1, hNotes + 1).setValue(newNotes);
      logAction("Atualização de Detalhes", tagSpool, `Estado: ${newStatus} | Notas atualizadas.`);
      break;
    }
  }
  return getInitialData();
}

function finishJointKiosk(jointId, tagSpool) {
  const ss = getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_2_WeldingMap_NDT');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const hJoint = headers.indexOf('ID_Junta');
  const hSpool = headers.indexOf('Tag_Spool');
  const hDate = headers.indexOf('Data_Soldadura');

  for (let i = 1; i < data.length; i++) {
    if (data[i][hJoint] == jointId && data[i][hSpool] == tagSpool) {
      sheet.getRange(i + 1, hDate + 1).setValue(new Date());
      logAction("Junta Soldada", tagSpool, `Junta ${jointId} terminada no Kiosk.`);
      break;
    }
  }
  return getInitialData();
}

function addInventoryMovement(movementData) {
  const ss = getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_4_Inventario_Logistica');
  const headers = sheet.getDataRange().getValues()[0];

  let newRow = new Array(headers.length).fill('');

  // Map custom frontend property names to sheet headers
  const mappedData = {
      Data_Movimento: new Date(),
      Tipo_Movimento: movementData.type,
      Codigo_SAP: movementData.sap,
      Quantidade: movementData.qty,
      Tag_Spool_Destino: movementData.type === 'SAIDA' ? movementData.spool : '',
      Heat_Number: movementData.heat
  };

  Object.keys(mappedData).forEach(key => {
    let index = headers.indexOf(key);
    if (index !== -1) {
      newRow[index] = mappedData[key];
    }
  });

  sheet.appendRow(newRow);
  logAction("Movimento Estoque", movementData.Codigo_SAP, `${movementData.Tipo_Movimento} - Qtd: ${movementData.Quantidade}`);
  return getInitialData();
}
