
const TEAMONE_VERSION = 'prototype-skeleton-1.0';


function gsTeamoneLicenseInfo() {
  return {
    version: TEAMONE_VERSION,
    docVersion: TEAMONE_VERSION,
    email: Session.getActiveUser().getEmail(),
    licExpiry: '',
    licStatus: 'OK',
    domainLic: '',
  };
}


function TEAMONE(cubeName, rowElements, columnElements, contextFilters, options, solveOrder) {
  if (solveOrder === undefined && typeof options === 'number') {
    solveOrder = options;
    options = '';
  }

  let sError = '';
  if (!('' + cubeName).trim()) {
    sError = 'ERROR: First argument must be a cube name, i.e.: Sales';
  } else if (Array.isArray(cubeName)) {
    sError = 'ERROR: First argument must be a cube name, not a cell range, i.e.: Sales';
  } else if (!rowElements && !columnElements && !contextFilters) {
    sError = 'ERROR: Provide at least one row, column or title element. You can skip arguments if needed, i.e. =TEAMONE("Sales",,B1)';
  } else if (Array.isArray(options)) {
    sError = 'ERROR: Options must be a text value like "S", "U", "SZ", not a cell range.';
  }

  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sSheetId = activeSheet.getSheetId();
  const cellPos = activeSheet.getActiveRange().getA1Notation();
  const propName = sSheetId + '__TEAMONE__' + cellPos;
  const docProperties = docPropsGet();

  if (sError) return sError;

  options = String(options || '').toUpperCase();
  const sFormula = JSON.stringify({
    cube: cubeName,
    rows: rowElements,
    columns: columnElements,
    titles: contextFilters,
    options: options,
    solveOrder: solveOrder || 0,
  });

  const key = sSheetId + '__VALUE__' + hashCode(sFormula);
  const sCellCache = docProperties.getProperty(key);
  if (!sCellCache) docProperties.setProperty(propName, sFormula);
  if (options.includes('X')) return 'Recalculation disabled (X flag)';
  return sCellCache ? JSON.parse(sCellCache).value : 'Recalculate in sidebar';
}


function MDX(cubeName, rowElements, columnElements, contextFilters, options){
  // Check cubeName is a single value string
  let sError = '';
  if(!(''+cubeName).trim()){
    sError = 'ERROR: First argument should be a cube name'
  }else if(Array.isArray(cubeName)){
    sError = 'ERROR: First argument should be a cube name'
  }
  options=String(options || '').toUpperCase();
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(), sSheetId = activeSheet.getSheetId();
  if(CacheService.getUserCache().get(sSheetId))return 'Recalculating...'
  const cellPos = activeSheet.getActiveRange().getA1Notation(), propName = sSheetId + '__MDX__' + cellPos, docProperties = docPropsGet();
  if(sError){
    
    return sError;
  }
  const sFormula = JSON.stringify({cube: cubeName, rows: rowElements, columns: columnElements, titles: contextFilters, options: options, solveOrder: 0}), sCellCache = docProperties.getProperty(sSheetId+'__VALUE__'+hashCode(sFormula));
  if(!sCellCache)docProperties.setProperty(propName, sFormula);
  if(options.includes('X')) return "Recalculation disabled (X flag)"
  return sCellCache?JSON.parse(sCellCache)['value']:'Recalculate in sidebar'
}


function gsTeamoneFormulaAdd(cube, rangeRows, rangeCols, rangeTitles, options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = sheet.getSelection().getCurrentCell();
  const nStartRow = cell.getValue() ? sheet.getLastRow() + 2 : cell.getRow();
  const nStartCol = cell.getColumn();
  const sStartCol = r1c1ToA1(nStartRow, nStartCol, true);
  const locale = ss.getSpreadsheetLocale();
  const argDelimiter = ['fr', 'de', 'es', 'it', 'pl', 'pt', 'ru', 'nl'].some(loc => locale.startsWith(loc)) ? ';' : ',';

  rangeRows = rangeRows && rangeRows.length ? rangeRows : [['Product A']];
  rangeCols = rangeCols && rangeCols.length ? rangeCols : [['Actual']];
  rangeTitles = rangeTitles || [];

  const nRowNodes = rangeRows[0].length;
  const nColNodes = rangeCols[0].length;
  const nTitleNodes = rangeTitles.length;
  const aData = [];

  aData.push([cube].concat(new Array(nColNodes + nRowNodes - 1).fill('')));
  rangeTitles.forEach(d => aData.push([d].concat(new Array(nColNodes + nRowNodes - 1).fill(''))));
  rangeCols.forEach(d => aData.push(new Array(nRowNodes).fill('').concat(d)));
  rangeRows.forEach(d => aData.push(d.concat(new Array(nColNodes).fill(''))));

  const sFormula = '=TEAMONE(' + sStartCol + nStartRow + argDelimiter
    + sStartCol + `${nStartRow + nTitleNodes + rangeCols.length + 1}:${r1c1ToA1(nStartRow + nTitleNodes + rangeCols.length + rangeRows.length, nStartCol + nRowNodes - 1)}` + argDelimiter
    + `${r1c1ToA1(nStartRow + nTitleNodes + 1, nStartCol + nRowNodes)}:${r1c1ToA1(nStartRow + nTitleNodes + rangeCols.length, nStartCol + nRowNodes + nColNodes - 1)}`
    + `${nTitleNodes > 0 ? (argDelimiter + sStartCol + (nStartRow + 1) + ':' + sStartCol + (nStartRow + nTitleNodes)) : argDelimiter}`
    + (options ? `${argDelimiter}"${options}"` : '')
    + ')';

  aData[nTitleNodes + rangeCols.length + 1][nRowNodes] = sFormula;
  gsDataPrint(sheet.getName(), aData, { keepData: true, startRow: nStartRow, startColumn: nStartCol });
  return { sheet: sheet.getName(), nStartRow, nStartCol, data: aData };
}

function gsInfo(scope, sheetId) {
  if (!scope) throw new Error('gsInfo(): scope is missing');

  
  const scopeSet = new Set(Array.isArray(scope) ? scope : [scope]);
  const has = (k) => scopeSet.has(k);

  const startTime = new Date().toISOString();
  const docProperties = PropertiesService.getDocumentProperties();
  const userProperties = PropertiesService.getUserProperties();
  const docPropKeys = docProperties.getKeys();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = ss.getSheets();
  const sheetIds = sheets.map(s => String(s.getSheetId()));
  const sheetIdSet = new Set(sheetIds);

  
  const sheetNameById = Object.fromEntries(
    sheets.map(s => [String(s.getSheetId()), s.getSheetName()])
  );

  // Resolve target sheet
  const targetSheetId = sheetId != null ? String(sheetId) : String(ss.getActiveSheet().getSheetId());
  const targetSheetName = sheetId != null ? sheetNameById[targetSheetId] : ss.getActiveSheet().getName();

  const sheet = sheetId != null ? ss.getSheetByName(targetSheetName) : ss.getActiveSheet();
  if (!sheet) {
    throw new Error(`Sheet not found: id=${targetSheetId}, name=${sheetNameById[targetSheetId]}`);
  }

  // Outputs
  let activeCell = {};
  let drillInfo;
  let explorationRange = null;
  let exploration = null;

  const formulas = [];
  const constantsSeen = new Set();

  // ----------------------------
  // ACTIVECELL
  // ----------------------------
  if (has('ACTIVECELL')) {
    const cell = sheet.getSelection().getCurrentCell();
    activeCell = {
      cell: cell.getA1Notation(),
      value: cell.getValue(),
      row: cell.getRow(),
      column: cell.getColumn(),
      values: []
    };

    const rangeList = sheet.getActiveRangeList();
    if (rangeList) {
      
      activeCell.values = rangeList.getRanges().flatMap(r => r.getValues());
    }
  }

  // ----------------------------
  // FORMULAS
  // ----------------------------
  if (has('FORMULAS')) {
   
    const valueKeysOnThisSheet = [];
    const expectedValueKeysFromFormulas = new Set();

    
    const formulaInCellCache = new Map();
    const getCellFormulaUpper = (a1) => {
      if (!formulaInCellCache.has(a1)) {
        
        const f = sheet.getRange(a1).getFormula() || '';
        formulaInCellCache.set(a1, f.toUpperCase());
      }
      return formulaInCellCache.get(a1);
    };

    
    const handlers = {
      TEAMONE: ({ key, sCell }) => {
        const cellFormula = getCellFormulaUpper(sCell);
        if (!cellFormula.includes('TEAMONE')) return { deleteKey: true };

        let sFormula = docProperties.getProperty(key);
        const sHash = hashCode(sFormula);
        let aFormulaRows, oFormula;

        
        if (sFormula.includes('__L__')) {
          try {
            const rhs = sFormula.split('__L__')[1];
            const rangeA1 = ((rhs.toUpperCase().split('TEAMONE(')[1]?.split(',')[1]) || '').trim();
            aFormulaRows = rangeA1 ? sheet.getRange(rangeA1).getValues() : undefined;

            if (!oFormula) oFormula = JSON.parse(sFormula);
            oFormula.rows = aFormulaRows;
            sFormula = JSON.stringify(oFormula);
          } catch (e) {
            throw new Error(`${e} ['L' option] Cannot evaluate row range in cell "${sCell}": ${sFormula.split('__L__')[1]}`);
          }
        }

        expectedValueKeysFromFormulas.add(`${targetSheetId}__VALUE__${sHash}`);

        let oCell = {}, value = '';
        if (has('VALUES')) {
          if (!oFormula) oFormula = JSON.parse(sFormula);
          const oCell = a1ToR1C1(sCell);

          let nRows = 1;
          let nCols = 1;

          if (aFormulaRows) nRows = Array.isArray(aFormulaRows) ? aFormulaRows.length : 1;
          else if (oFormula.rows) nRows = Array.isArray(oFormula.rows) ? oFormula.rows.length : 1;

          if (oFormula.columns) {
            nCols = Array.isArray(oFormula.columns)
              ? (Array.isArray(oFormula.columns[0]) ? oFormula.columns[0].length : 1)
              : 1;
          }
          value = sheet.getRange(oCell.row, oCell.column, nRows, nCols).getValues();
        }
        if (has('DRILL_INFO') && !drillInfo) {
          oCell = a1ToR1C1(sCell);
          if (!oFormula) oFormula = JSON.parse(sFormula);
          const formulaRange = sheet.getRange(oCell.row, oCell.column, Array.isArray(oFormula.rows) ? oFormula.rows.length : 1, Array.isArray(oFormula.columns) ? oFormula.columns[0].length : 1);
          const [rngRow, rngRows, rngCol, rngCols] = [
            formulaRange.getRow(),
            formulaRange.getNumRows(),
            formulaRange.getColumn(),
            formulaRange.getNumColumns()
          ];
          
          if (!(activeCell.row < rngRow
            || activeCell.row > rngRow + rngRows - 1
            || activeCell.column < rngCol
            || activeCell.column > rngCol + rngCols - 1
          )) {
            drillInfo = {cell: sCell, cube: oFormula.cube, source: 'TEAMONE', formula: sFormula};  
            
            
            if (Array.isArray(oFormula.rows)) {
              drillInfo.nRowDims = oFormula.rows.length;
              drillInfo.rows = fillLastRow(oFormula.rows.slice(0, activeCell.row - rngRow + 1))
            } else {
              drillInfo.nRowDims = 1;
              drillInfo.rows = [[oFormula.rows]]
            }

            
            if (Array.isArray(oFormula.columns)) {
              drillInfo.iCol = activeCell.column - rngCol;
              drillInfo.columns = fillLastColumn(oFormula.columns.map(r=>r.slice(0, drillInfo.iCol + 1)));
            } else {
              drillInfo.nColDims = 1;
              drillInfo.columns = [[oFormula.columns]];
            }
            drillInfo.titles = Array.isArray(oFormula.titles) ? oFormula.titles.flat(Infinity).filter(v => v != null && v !== '') : [oFormula.titles];
          }
        }
        formulas.push({ cell: sCell, cellInfo: oCell, formula: sFormula, value, name: 'TEAMONE' });
        return {};
      },

      ATTR: ({ key, sCell }) => {
        const cellFormula = getCellFormulaUpper(sCell);
        if (!cellFormula.includes('ATTR')) return { deleteKey: true };

        const sFormula = docProperties.getProperty(key);
        expectedValueKeysFromFormulas.add(`${targetSheetId}__VALUE__${hashCode(sFormula)}`);

        formulas.push({ cell: sCell, formula: sFormula, value: '', name: 'ATTR' });
        return {};
      },

      SUBSET: ({ key, sCell }) => {
        const cellFormula = getCellFormulaUpper(sCell);
        if (!cellFormula.includes('SUBSET')) return { deleteKey: true };

        const sFormula = docProperties.getProperty(key);
        expectedValueKeysFromFormulas.add(`${targetSheetId}__VALUE__${hashCode(sFormula)}`);

        formulas.push({ cell: sCell, formula: sFormula, value: [[]], name: 'SUBSET' });
        return {};
      },

      MDX: ({ key, sCell }) => {
        const cellFormula = getCellFormulaUpper(sCell);
        if (!cellFormula.includes('MDX')) return { deleteKey: true };

        const sFormula = docProperties.getProperty(key);
        const valueKey = `${targetSheetId}__VALUE__${hashCode(sFormula)}`;
        expectedValueKeysFromFormulas.add(valueKey);

        let oCell = {}, value = '';
        const sRangeName = `T${targetSheetId}${sCell}__MDX`;
        if (has('VALUES')) {
          const oFormula = JSON.parse(sFormula);
          oCell = a1ToR1C1(sCell);
          value = ss.getRangeByName(sRangeName).getValues();
          
          if(!(oFormula.rows || '').includes(' * ')) {
            const nCols = Array.isArray((oFormula.columns)) ? oFormula.columns.length : (oFormula.columns || '').split(' * ').length;
            const aRowElm = Array.isArray(oFormula.rows) ? oFormula.rows : [[oFormula.rows || '']]
          }
          
        }
        if (has('DRILL_INFO') && !drillInfo) {
          oCell = a1ToR1C1(sCell);
          const oFormula = JSON.parse(sFormula);
          const formulaRange = ss.getRangeByName(sRangeName);
          let staticCols = true;
          
          const [rngRow, rngRows, rngCol, rngCols] = [
            formulaRange.getRow(),
            formulaRange.getNumRows(),
            formulaRange.getColumn(),
            formulaRange.getNumColumns()
          ];
          
          if (!(activeCell.row < rngRow
            || activeCell.row > rngRow + rngRows - 1
            || activeCell.column < rngCol
            || activeCell.column > rngCol + rngCols - 1
          )) {
            drillInfo = {cell: sCell, cube: oFormula.cube, source: 'MDX', formula: sFormula, rowDims: [], colDims: []};
            
            drillInfo.rowDims = (oFormula.rows || '').split(' * ').map( (mdx)=>{ const aExp = mdx.split('^'); return String(aExp[0]).trim() + '^' + String(aExp.length>2 ? aExp[1] : aExp[0]).trim() + '^' } )
            drillInfo.nRowDims = drillInfo.rowDims.length;
            if (Array.isArray(oFormula.columns)) {
              drillInfo.iCol = activeCell.column - rngCol - drillInfo.nRowDims;
              drillInfo.columns = fillLastColumn(oFormula.columns.map(r=>r.slice(0, drillInfo.iCol + 1)));
              drillInfo.nColDims = oFormula.columns.length;
            } else {

              staticCols = false;
              drillInfo.colDims = (oFormula.columns || '').split(' * ').map( (mdx)=>{ const aExp = mdx.split('^'); return String(aExp[0]).trim() + '^' + String(aExp.length>2 ? aExp[1] : aExp[0]).trim() + '^' } )
              drillInfo.nColDims = drillInfo.colDims.length;

              drillInfo.columns = fillLastColumn(sheet.getRange(`${r1c1ToA1(rngRow, rngCol+drillInfo.nRowDims)}:${r1c1ToA1(rngRow + drillInfo.nColDims - 1, activeCell.column)}`).getValues(), drillInfo.colDims);
            }
            if (activeCell.column < rngCol + drillInfo.nColDims || activeCell.row < rngRow + (staticCols ? 0 : drillInfo.nRowDims)) throw 'Please select a data cell (value) to drill through, not a row or column header'

            
            drillInfo.rows = fillLastRow(sheet.getRange(`${staticCols ? sCell : r1c1ToA1(rngRow + drillInfo.nColDims, rngCol)}:${r1c1ToA1(activeCell.row, rngCol + drillInfo.nRowDims - 1)}`).getValues(), drillInfo.rowDims);
            drillInfo.titles = drillInfo.titles = Array.isArray(oFormula.titles) ? oFormula.titles.flat(Infinity).filter(v => v != null && v !== '') : [oFormula.titles];
          }
        }
        formulas.push({ cell: sCell, cellInfo: oCell, formula: sFormula, value, name: 'MDX' });
        return {};
      },

      USERNAME: ({ key }) => {
        if (constantsSeen.has('USERNAME')) return {};
        expectedValueKeysFromFormulas.add(`${targetSheetId}__VALUE__USERNAME`);
        formulas.push({ cell: 'SHEET', formula: docProperties.getProperty(key), value: [[]], name: 'USERNAME' });
        constantsSeen.add('USERNAME');
        return {};
      },

      UPDATED: ({ key }) => {
        if (constantsSeen.has('UPDATED')) return {};
        expectedValueKeysFromFormulas.add(`${targetSheetId}__VALUE__UPDATED`);
        formulas.push({ cell: 'SHEET', formula: docProperties.getProperty(key), value: [[]], name: 'UPDATED' });
        constantsSeen.add('UPDATED');
        return {};
      }
    };

    const formulaNamesAll = new Set(['VALUE', 'TEAMONE', 'ATTR', 'SUBSET', 'MDX', 'USERNAME', 'UPDATED']);

    for (const key of docPropKeys) {
      if (!key.includes('__')) continue;
      const [sSheetIdThis, sFormulaName, sCell] = key.split('__');
      if (!formulaNamesAll.has(sFormulaName)) continue;

      
      if (sFormulaName === 'VALUE') {
        if (sSheetIdThis === targetSheetId) valueKeysOnThisSheet.push(key);
        else if (!sheetIdSet.has(sSheetIdThis)) docProperties.deleteProperty(key);
        continue;
      }

     
      if (!sheetIdSet.has(sSheetIdThis)) {
        docProperties.deleteProperty(key);
        continue;
      }

     
      if (sSheetIdThis !== targetSheetId) continue;

      const handler = handlers[sFormulaName];
      if (!handler) continue;

      const { deleteKey } = handler({ key, sCell, sSheetIdThis });
      if (deleteKey) docProperties.deleteProperty(key);
    }

   
    for (const k of valueKeysOnThisSheet) {
      if (!expectedValueKeysFromFormulas.has(k)) docProperties.deleteProperty(k);
    }
  }

  // ----------------------------
  // TEAMONESHEETS
  // ----------------------------
  let teamoneSheets = [];
  if (has('TEAMONESHEETS')) {
    const sheetNameLower = sheets.map(s => s.getSheetName().toLowerCase());
    const sheetIdByLowerName = Object.fromEntries(
      sheets.map(s => [s.getSheetName().toLowerCase(), s.getSheetId()])
    );

    const sheetList = {}; 
    const explorations = JSON.parse(docProperties.getProperty('EXPLORATION') ?? '[]');

    for (const e of explorations) {
      const eId = String(e.id);
      if (sheetIdSet.has(eId)) {
        sheetList[eId] = { id: eId, sheet: sheetNameById[eId], exploration: true, formula: [] };
      } else {
        const lower = String(e.sheet ?? '').toLowerCase();
        if (sheetNameLower.includes(lower)) {
          const fixedId = String(sheetIdByLowerName[lower]);
          e.id = fixedId;
          sheetList[fixedId] = { id: fixedId, sheet: e.sheet, exploration: true, formula: [] };
        }
      }
    }

    const formulaNames = new Set(['TEAMONE', 'ATTR', 'SUBSET', 'MDX']);

    const protectedKeys = new Set(['EXPLORATION', 'MODAL_STATE', 'TEAMONE_STATE', 'DOC_CONFIG', 'TEAMONE_CONFIG']);

    for (const key of docPropKeys) {
      
      if (!key.includes('__') || protectedKeys.has(key)) continue;

      const [sSheetIdThis, sFunctionName, sCell] = key.split('__');
      if (!formulaNames.has(sFunctionName)) continue;
      if (!sheetIdSet.has(sSheetIdThis)) continue;

      if (!sheetList[sSheetIdThis]) {
        sheetList[sSheetIdThis] = {
          id: sSheetIdThis,
          sheet: sheetNameById[sSheetIdThis],
          exploration: false,
          formula: []
        };
      }
      sheetList[sSheetIdThis].formula.push(sCell);
    }

    // Add .TEAMONE sheet (if not already added)
    if (sheetNameLower.includes('.teamone')) {
      const id = String(sheetIdByLowerName['.teamone']);
      if (id && !sheetList[id]) sheetList[id] = { id, sheet: '.TEAMONE', exploration: false, formula: [] };
    }

    teamoneSheets = Object.values(sheetList);
  }

  // ----------------------------
  // SHEETS
  // ----------------------------
  const sheetsMap = has('SHEETS') ? sheetNameById : {};

  return {
    startTime,
    endTime: new Date().toISOString(),
    scope: Array.from(scopeSet),
    sheet: sheet.getName(),
    id: String(sheet.getSheetId()),
    ACTIVECELL: activeCell,
    TEAMONESHEETS: teamoneSheets,
    FORMULAS: formulas,
    DRILL_INFO: drillInfo,
    EXPLORATION: exploration,
    EXPLORATION_RANGE: explorationRange,
    CLIPBOARD: has('CLIPBOARD') ? userProperties.getProperty('CLIPBOARD') : null,
    SHEETS: sheetsMap,
    PROPS: has('PROPS') ? docProperties.getProperties() : null
  };
}


/**
 * Generic helper: insert a 2D array into the active sheet, starting at startCell.
 * (Not currently called by the assistant flow - gsTeamoneFormulaAdd() is used
 * for that - kept here in case a future feature needs a plain range insert.)
 */
function insertAssistantRange(startCell, values) {
  if (!Array.isArray(values) || !values.length || !Array.isArray(values[0])) {
    throw new Error('insertAssistantRange(): values must be a non-empty 2D array.');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(startCell || 'A1');
  const numRows = values.length;
  const numCols = Math.max.apply(null, values.map(row => row.length));
  const normalized = values.map(row => {
    const copy = row.slice();
    while (copy.length < numCols) copy.push('');
    return copy;
  });

  sheet.getRange(range.getRow(), range.getColumn(), numRows, numCols).setValues(normalized);
  return { sheet: sheet.getName(), startCell: range.getA1Notation(), rows: numRows, columns: numCols };
}

/**
 * Activate sheet by name or ID.
 */
function gsSheetActivate(options) {
  options = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;
  if (options.id) ss.getSheets().forEach(s => { if (s.getSheetId() == options.id) sheet = s; });
  if (!sheet && options.sheet) sheet = ss.getSheetByName(options.sheet);
  if (!sheet && options.createIfNotFound && options.sheet) sheet = ss.insertSheet(options.sheet);
  if (!sheet) throw new Error('Sheet not found: ' + (options.sheet || options.id));
  sheet.activate();
  if (options.cell) sheet.setCurrentCell(sheet.getRange(options.cell));
  return { sheet: sheet.getName(), id: sheet.getSheetId() };
}

function gsTeamoneStateGet() {
  return PropertiesService.getDocumentProperties().getProperty('TEAMONE_PROTOTYPE_STATE') || '{}';
}

function gsTeamoneStateSave(state) {
  PropertiesService.getDocumentProperties().setProperty('TEAMONE_PROTOTYPE_STATE', String(state || '{}'));
  return { status: 'OK' };
}

function gsDataPrint(sheetName, data, options) {
  options = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);
  if (!Array.isArray(data) || !data.length) return { sheet: sheet.getName(), rows: 0, columns: 0 };

  const startRow = options.startRow || 1;
  const startColumn = options.startColumn || 1;
  const numRows = data.length;
  const numCols = Math.max.apply(null, data.map(row => row.length));
  const normalized = data.map(row => {
    const copy = row.slice();
    while (copy.length < numCols) copy.push('');
    return copy;
  });

  sheet.getRange(startRow, startColumn, numRows, numCols).setValues(normalized);
  if (options.autofitCols) sheet.autoResizeColumns(startColumn, numCols);
  return { sheet: sheet.getName(), startRow, startColumn, rows: numRows, columns: numCols };
}

function r1c1ToA1(row, column, columnOnly) {
  let temp = '';
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return columnOnly ? letter : letter + row;
}

function hashCode(str) {
  str = String(str || '');
  let hash = 0;
  if (str.length === 0) return hash;
  for (let i = 0; i < str.length; i++) {
    const chr = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + chr;
    hash |= 0;
  }
  return String(hash);
}

function docPropsGet() {
  return PropertiesService.getDocumentProperties();
}

function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

function onInstall() {
  onOpen();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TeamOne')
    .addItem('TeamOne Sidebar', 'menuShowTeamoneSidebar')
    .addToUi();
}

function menuShowTeamoneSidebar() {
  const html = HtmlService.createTemplateFromFile('core');
  html.page = { file: 'sidebar', displayAsSidebar: true, head: '', js: '' };
  SpreadsheetApp.getUi().showSidebar(html.evaluate().setTitle('TeamOne'));
}

function getBackendUrl() {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('TEAMONE_BACKEND_URL');
  if (!url) {
    throw new Error('TEAMONE_BACKEND_URL not set in Script Properties (Project Settings > Script Properties).');
  }
  return url;
}

function getOrCreateAssistantSessionId() {
  const props = PropertiesService.getUserProperties();
  let id = props.getProperty('TEAMONE_ASSISTANT_SESSION_ID');
  if (!id) {
    id = Utilities.getUuid();
    props.setProperty('TEAMONE_ASSISTANT_SESSION_ID', id);
  }
  return id;
}
