// test

const COLUMNS_NAMES = {
    "cost": ["abrechnungsbetrag in fondswährung", "kundenendbetrag eur"],
    "currencies": ["fondswährung", "währung"],
    "dates": ["abrechnungstag", "buchungsdatum"],
    "ISINs": ["isin"],
    "kurses": ["abrechnungspreis", "kurs"],
    "quantities": ["stücke/nom.", "anteile"],
    "typesOfDeals": ["transaktion", "geschäftsart"],
    "clearCost": ["kurswert eur"]
}

const DEAL_TYPE = {
    "buy": ["kauf aus sparplan",
        "kauf",
        "ertrag wiederanlage",
        "tausch (kauf)",
        "fondsmerge steuerpflichtig (zugang)",
        "tausch gesamt (kauf)",
        "fondsmerge steuerneutral (zugang)"],
    "ignore": ["steuererstattung",
        "ertrag",
        "ertrag auszahlung",
        "vorabpauschale abrechnung lastschrift",
        "steuererstattung",
        "steuerforderung",
        "vorabpauschale abrechnung",
        "entgeltbelastung",
        "delta-korrektur abgang",
        "storno ertrag ohne wiederanlage"],
    "sell": ["tausch (verkauf)",
        "verkauf",
        "tausch (verkauf)",
        "fondsmerge steuerpflichtig (abgang)",
        "vorabpauschale verkauf",
        "storno verkauf",
        "tausch gesamt (verkauf)",
        "fondsmerge steuerneutral (abgang)"]
}

const TAXES_SHEET_NAME = ["taxes", "steuern_isin"];

const TAXES_COLUMNS_NAMES = {
    "basiszins": ["basiszins"],
    "basiszinsanteil": ["basiszinssatz anteil"],
    "basiszinssatz": ["basiszinssatz"],
    "bbzinsen": ["bundesbank zinsen"],
    "kapital": ["kapitalertragsteuer"],
    "solidar": ["solidaritätszuschlag"],
    "tax": ["steuer", "tax"],
    "year": ["jahr", "year"]
}


const ISIN_TAXES_SHEET_NAME = "isin_taxes";

class ColumnWriter {
    constructor(sheet, column, row) {
        this.sheet = sheet;
        this.column = column;
        this.row = row;
    }

    write(data){
        this.sheet.getRange(this.row, this.column).setValue(data);
        this.column++;
        return this.column - 1;
    }

    setFormula(formula){
        this.sheet.getRange(this.row, this.column).setFormula(formula);
        this.column++;
        return this.column - 1;
    }

    nextRow(){
        this.column = 1;
        this.row++;
    }
}

/**
 * @OnlyCurrentDoc
 */
function onOpen() {
    SpreadsheetApp.getUi().createMenu('Custom Menu').addItem('Evaluate FIFO', 'menuItemEval').addToUi();
}

function menuItemEval() {
    // need to scan all data, and ask needable data for taxes, only once, then save it in PropertiesService
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getActiveSheet();
    let data = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues().flat();

    let columns = { "dates": undefined, "ISINs": undefined, "typesOfDeals": undefined, "quantities": undefined, "cost": undefined, "currencies": undefined, "kurses": undefined };

    for (let i = 0; i < data.length; i++) {
        if (data[i] === undefined || typeof data[i] !== "string") {
            continue;
        }
        let lower_data = data[i].toLowerCase();

        if (COLUMNS_NAMES.dates.includes(lower_data)) {
            columns.dates = i;
        } else if (COLUMNS_NAMES.ISINs.includes(lower_data)) {
            columns.ISINs = i;
        } else if (COLUMNS_NAMES.typesOfDeals.includes(lower_data)) {
            columns.typesOfDeals = i;
        } else if (COLUMNS_NAMES.quantities.includes(lower_data)) {
            columns.quantities = i;
        } else if (COLUMNS_NAMES.cost.includes(lower_data)) {
            columns.cost = parseFloat(i);
        } else if (COLUMNS_NAMES.currencies.includes(lower_data)) {
            columns.currencies = i;
        } else if (COLUMNS_NAMES.kurses.includes(lower_data)) {
            columns.kurses = i;
        } else if (COLUMNS_NAMES.clearCost.includes(lower_data)) {
            columns.clearCost = i;
        }
    }

    // check all columns are found and return missing column in alert
    let outData = {
        "dates": undefined,
        "ISINs": undefined,
        "typesOfDeals": undefined,
        "quantities": undefined,
        "cost": undefined,
        "currencies": undefined,
        "kurses": undefined,
        "clearCost" : undefined
    }

    let keys = Object.keys(columns);

    for (let i = 0; i < keys.length; i++) {
        if (columns[keys[i]] === undefined) {
            throw "column " + keys[i] + " is missing";
        }

        outData[keys[i]] = sheet.getRange(2, columns[keys[i]] + 1, sheet.getLastRow() - 1, 1).getValues().flat();
    }

    // check if all input arguments have the same length
    let length = outData[keys[0]].length;
    for (let i = 1; i < keys.length; i++) {
        if (outData[keys[i]].length !== length) {
            //SpreadsheetApp.getUi().alert("input arguments have different length");
            throw "input arguments have different length";
        }
    }

    
    let lastColInd = sheet.getLastColumn();
    if (sheet.getRange(1, lastColInd - 6, 1, 7).getValues().toString() === [["Nettogewinne", "Gewinne/Verluste", "Teilfreistellung", "Gewinne/Verluste ohne Teilfreistellung", "Steuerabzug, eur", "Kapitalertragsteuer", "Solidaritätszuschlag"]].toString()) {
        // clear old results
        sheet.getRange(1, lastColInd - 6, sheet.getLastRow(), 7).clearContent();
        evalFifo(outData, lastColInd - 6);
    } else {
        evalFifo(outData, lastColInd + 1);
    }
}

let convertStrToDate = (str) => {
    if (str instanceof Date) {
        return str;
    }
    if (str.includes(".")) {
        let buf = str.split(".");
        return new Date(buf[2], buf[1] - 1, buf[0]);
    }
    else if (str.includes("-")) {
        let buf = str.split("-");
        return new Date(buf[0], buf[1] - 1, buf[2]);
    }
    return Date.parse(str);
}

let evalFifo = (inData, columnToPasteRes) => {
    // getting all data from input aguments in an array of maps
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let data = [];
    let currentBalace = new Map();
    let resultsForNewSheet = new Map();
    let averageKurs = new Map();
    let yearStockBalance = new Map();
    let ISINsUnique = new Set();

    for (let i = 0; i < inData.dates.length; i++) {
        let bufMap = new Map();

        bufMap.set("addressToWrite", i + 2)

        let bufDate = convertStrToDate(inData.dates[i]);

        if (isNaN(bufDate)) {
            //SpreadsheetApp.getUi().alert("element in dates is not a date (row: " + (i + 1) + ")");
            throw "element in dates is not a date (row: " + (i + 2) + ")";
        }

        bufMap.set("date", bufDate);
        bufMap.set("ISIN", inData.ISINs[i]);

        let dealTypeLower = inData.typesOfDeals[i].toLowerCase();
        // figure out type of deal
        if (DEAL_TYPE.buy.includes(dealTypeLower)) {
            bufMap.set("typeOfDeal", "Kauf");
        } else if (DEAL_TYPE.sell.includes(dealTypeLower)) {
            bufMap.set("typeOfDeal", "Verkauf");
        } else if (!(DEAL_TYPE.ignore.includes(dealTypeLower))) {
            //SpreadsheetApp.getUi().alert("element in typesOfDeals is not a valid type (row: " + (i + 1) + ")");
            throw "element in typesOfDeals is not a valid type (row: " + (i + 2) + ")";
        }

        if (isNaN(inData.quantities[i])) {
            //SpreadsheetApp.getUi().alert("element in quantities is not a number (row: " + (i + 1) + ")");
            throw "element in quantities is not a number (row: " + (i + 2) + ")";
        }
        if (bufMap.get("typeOfDeal") === "Verkauf" && inData.quantities[i] > 0) {
            bufMap.set("quantity", -inData.quantities[i]);
        } else {
            bufMap.set("quantity", inData.quantities[i]);
        }

        if (isNaN(inData.cost[i])) {
            throw "element in cost is not a number (row: " + (i + 2) + ")";
        } else if (inData.cost[i] === undefined || inData.cost[i] === null || inData.cost[i] === "") {
            throw "zero value in cost (row: " + (i + 2) + ")";
        }

        if (bufMap.get("typeOfDeal") === "Verkauf" && inData.cost[i] < 0) {
            bufMap.set("cost", -inData.cost[i]);
        } else {
            bufMap.set("cost", inData.cost[i]);
        }

        if (bufMap.get("typeOfDeal") === "Verkauf" && inData.clearCost[i] < 0) {
            bufMap.set("clearCost", -inData.clearCost[i]);
        } else {
            bufMap.set("clearCost", inData.clearCost[i]);
        }

        bufMap.set("kurs", Math.abs(inData.cost[i] / bufMap.get("quantity")));
        bufMap.set("clearKurs", Math.abs(inData.clearCost[i] / bufMap.get("quantity")));

        bufMap.set("currency", inData.currencies[i]);

        data.push(bufMap);

        let key = inData.ISINs[i] + " " + inData.currencies[i]
        currentBalace.set(key, 0);

        if (resultsForNewSheet.get(key) === undefined) {
            resultsForNewSheet.set(key, new Map())
        }
        resultsForNewSheet.get(key).set(bufDate.getFullYear(), 0);

        if (yearStockBalance.get(key) === undefined) {
            yearStockBalance.set(key, new Map());
        }
        yearStockBalance.get(key).set(bufDate.getFullYear(), 0);

        ISINsUnique.add(inData.ISINs[i]);
    }

    // sort elemnts by dateinTypesOfDeals
    data.sort((a, b) => a.get("date") - b.get("date"));


    // create array of years from data
    let years = [];
    for (let i = 0; i < data.length; i++) {
        if (years.indexOf(data[i].get("date").getFullYear()) === -1) {
            years.push(data[i].get("date").getFullYear());
        }
    }

    let taxes = new Map();
    // search for sheet name
    let leftYears = [...years];
    let taxSheet = null;
    for (i = 0; i < TAXES_SHEET_NAME.length; i++) {
        taxSheet = ss.getSheetByName(TAXES_SHEET_NAME[i])
        if (taxSheet !== null) {
            break;
        }
    }

    if (taxSheet === null) {
        throw "sheets " + TAXES_SHEET_NAME + " not found";
    }

    // check all years are in sheet
    let columnsAdresses = new Map();
    let leftColumns = [...Object.keys(TAXES_COLUMNS_NAMES)];
    for (let i = 1; i < taxSheet.getLastColumn() + 1; i++) {
        let buf = taxSheet.getRange(1, i).getValue();
        for (let j = 0; j < Object.keys(TAXES_COLUMNS_NAMES).length; j++) {
            if (TAXES_COLUMNS_NAMES[Object.keys(TAXES_COLUMNS_NAMES)[j]].includes(buf.toLowerCase())) {
                columnsAdresses.set(Object.keys(TAXES_COLUMNS_NAMES)[j], i);
                leftColumns.splice(leftColumns.indexOf(Object.keys(TAXES_COLUMNS_NAMES)[j]), 1);
            }
        }
    }

    if (leftColumns.length > 0) {
        throw "sheet " + taxSheet.getName() + " does not contain all columns, missing columns: " + leftColumns.join(', ');
    }

    let yearsRows = new Map();
    for (let i = 2; i < taxSheet.getLastRow() + 1; i++) {
        let year = Number(taxSheet.getRange(i, columnsAdresses.get("year")).getValue());

        if (isNaN(year)) {
            throw "sheet " + taxSheet.getName() + " contains not a number in row " + (i);
        }

        if (leftYears.indexOf(year) !== -1) {
            leftYears.splice(leftYears.indexOf(year), 1);
        }

        yearsRows.set(year, i);
    }

    if (leftYears.length > 0) {
        throw "sheet " + taxSheet.getName() + " does not contain all years, missing years: " + leftYears.join(', ');
    }

    // write all adresses in map
    for (let i = 0; i < years.length; i++) {
        taxes.set(years[i], {
            "tax": `'` + taxSheet.getName() + `'!` + columnToLetter(columnsAdresses.get("tax")) + yearsRows.get(years[i]),
            "basiszins": `'` + taxSheet.getName() + `'!` + columnToLetter(columnsAdresses.get("basiszins")) + yearsRows.get(years[i]),
            "solidar": `'` + taxSheet.getName() + `'!` + columnToLetter(columnsAdresses.get("solidar")) + yearsRows.get(years[i]),
            "kapital": `'` + taxSheet.getName() + `'!` + columnToLetter(columnsAdresses.get("kapital")) + yearsRows.get(years[i]),
            "bbzinsen": `'` + taxSheet.getName() + `'!` + columnToLetter(columnsAdresses.get("bbzinsen")) + yearsRows.get(years[i]),
            "basiszinsAnteil": `'` + taxSheet.getName() + `'!` + columnToLetter(columnsAdresses.get("basiszinsAnteil")) + yearsRows.get(years[i]),
            "basiszinssatz": `'` + taxSheet.getName() + `'!` + columnToLetter(columnsAdresses.get("basiszinssatz")) + yearsRows.get(years[i])
        });
    }

    // collect ISIN taxes
    let ISINtaxes = new Map();
    let leftISINs = [...ISINsUnique];

    // searching for columns in sheet
    let ISINColumns = { "ISIN": undefined, "teilfreistellung": undefined };
    for (i = 1; i < taxSheet.getLastColumn() + 1; i++) {
        let buf = taxSheet.getRange(1, i).getValue();
        if (buf.toLowerCase() === "isin") {
            ISINColumns.ISIN = i;
        } else if (buf.toLowerCase() === "teilfreistellung") {
            ISINColumns.teilfreistellung = i;
        }
    }

    // check undefined columns
    for (i = 0; i < Object.keys(ISINColumns).length; i++) {
        if (ISINColumns[Object.keys(ISINColumns)[i]] === undefined) {
            throw "sheet " + taxSheet.getName() + " does not contain column " + Object.keys(ISINColumns)[i];
        }
    }

    for (let j = 2; j <= taxSheet.getLastRow(); j++) {
        let ISIN = taxSheet.getRange(j, ISINColumns.ISIN).getValue();
        let teilfreistellung = `'` + taxSheet.getName() + `'!` + columnToLetter(ISINColumns.teilfreistellung) + j;

        // check if ISIN is already was in sheet
        if (ISINtaxes.get(ISIN) !== undefined) {
            throw "ISIN " + ISIN + " was already in sheet isin_taxes (second appearance in row: " + j + ")";
        }

        leftISINs.splice(leftISINs.indexOf(ISIN), 1);

        ISINtaxes.set(ISIN, teilfreistellung);
    }


    // check that all ISINs found
    if (leftISINs.length > 0) {
        throw "sheet " + ISIN_TAXES_SHEET_NAME + " does not contain all ISINs, missing ISINs: " + leftISINs.join(', ');
    }

    let ISINrecords = new Map();

    const sheet = ss.getActiveSheet();

    // write headers
    let headers = ["Nettogewinne", "Gewinne/Verluste", "Teilfreistellung", "Gewinne/Verluste ohne Teilfreistellung", "Steuerabzug, eur", "Kapitalertragsteuer", "Solidaritätszuschlag"];
    for (let i = columnToPasteRes; i < headers.length + columnToPasteRes; i++){
        sheet.getRange(1, i).setValue(headers[i - columnToPasteRes]);
    }


    for (let i = 0; i < data.length; i++) {
        let nowKey = data[i].get("ISIN") + " " + data[i].get("currency");

        // average kurs for only unrealized
        if (ISINrecords.get(nowKey) !== undefined) {
            if (i + 1 < data.length && data[i].get("date").getFullYear() !== data[i + 1].get("date").getFullYear()) {
                let bufRecords = ISINrecords.get(nowKey);
                let bufAverageKurs = [];
                for (let j = 0; j < bufRecords.length; j++) {
                    bufAverageKurs.push(bufRecords[j].get("kurs"));
                }
                if (averageKurs.get(nowKey) === undefined) {
                    averageKurs.set(nowKey, new Map());
                }
                averageKurs.get(nowKey).set(data[i].get("date").getFullYear(), bufAverageKurs);
            }
        }

        if (currentBalace.get(nowKey) === 0) {
            // first deal with this ISIN
            //Logger.log("first deal with this ISIN")
            currentBalace.set(nowKey, data[i].get("quantity"));
            ISINrecords.set(nowKey, [data[i]]);
        } else if (currentBalace.get(nowKey) > 0 && data[i].get("typeOfDeal") === "Kauf") {
            // just add to current balance, no result
            //Logger.log("just add to current balance, no result")
            currentBalace.set(nowKey, currentBalace.get(nowKey) + data[i].get("quantity"));
            ISINrecords.get(nowKey).push(data[i]);
        } else if (currentBalace.get(nowKey) > 0 && data[i].get("typeOfDeal") === "Verkauf") {
            if (currentBalace.get(nowKey) < data[i].get("quantity")) {
                throw "not enough quantity in row " + (i + 2);
            }

            // need to evaluate result
            //Logger.log("need to evaluate result Verkauf")
            let records = ISINrecords.get(nowKey);
            let lastDeal = data[i];
            let balance = currentBalace.get(nowKey);
            let [bufRecords, bufBalance, bufResult] = processDeal(records, balance, lastDeal);
            ISINrecords.set(nowKey, bufRecords);
            currentBalace.set(nowKey, bufBalance);

            //Logger.log(ISINtaxes.get(nowKey.split(" ")[0]))
            //Logger.log("write to row %s result %s", data[i].get("addresToWrite"), bufResult)
            addDealRes(sheet, columnToPasteRes, data[i].get("addressToWrite"), bufResult, taxes, ISINtaxes.get(nowKey.split(" ")[0]), data[i].get("date").getFullYear());
            resultsForNewSheet.get(nowKey).set(data[i].get("date").getFullYear(), resultsForNewSheet.get(nowKey).get(data[i].get("date").getFullYear()) + bufResult);
        } else if (currentBalace.get(nowKey) < 0 && data[i].get("typeOfDeal") === "Kauf") {
            throw "short detected in row " + (i + 2);
        } else if (currentBalace.get(nowKey) < 0) {
            // just add to current balance, no result
            //Logger.log("just add to current balance, no result")
            currentBalace.set(nowKey, currentBalace.get(nowKey) + data[i].get("quantity"));
            ISINrecords.get(nowKey).push(data[i]);
        }
        yearStockBalance.get(nowKey).set(data[i].get("date").getFullYear(), currentBalace.get(nowKey));
    }

    // add rest to averageKurs
    let lastRecYear = data[data.length - 1].get("date").getFullYear();
    let ISINsKeys = Array.from(ISINrecords.keys());
    for (let i = 0; i < ISINsKeys.length; i++) {
        let bufRecords = ISINrecords.get(ISINsKeys[i]);
        let bufAverageKurs = [];
        for (let j = 0; j < bufRecords.length; j++) {
            bufAverageKurs.push(bufRecords[j].get("kurs"));
        }
        if (averageKurs.get(ISINsKeys[i]) === undefined) {
            averageKurs.set(ISINsKeys[i], new Map());
        }
        averageKurs.get(ISINsKeys[i]).set(lastRecYear, bufAverageKurs);
    }


    // clear formatting
    sheet.getRange(2, columnToPasteRes, sheet.getLastRow() - 1, headers.length).clearFormat();

    // two decimal places for all cells
    sheet.getRange(2, columnToPasteRes, sheet.getLastRow() - 1, headers.length).setNumberFormat("#,##0.00");

    // for the most useful results 

    let rules = [
        SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#36F566'),
        SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground('#C70000'),
    ]

    // apply formatting
    let allRules = sheet.getConditionalFormatRules();
    for (let i = 0; i < rules.length; i++) {
        allRules.push(rules[i].setRanges([sheet.getRange(2, columnToPasteRes, sheet.getLastRow() - 1, 1)]).build());
    }
    sheet.setConditionalFormatRules(allRules);

    rules = [
        SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#58D68D'),
        SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground('#EC7063'),
    ]

    // apply formatting
    allRules = sheet.getConditionalFormatRules();
    for (let i = 0; i < rules.length; i++) {
        allRules.push(rules[i].setRanges([sheet.getRange(2, columnToPasteRes + 1, sheet.getLastRow() - 1, headers.length - 1)]).build());
    }
    sheet.setConditionalFormatRules(allRules);

    // resize edited columns to fit content
    sheet.autoResizeColumns(columnToPasteRes, headers.length);

    // write header
    headers = ['ISIN', 'Symbol (von https://finance.yahoo.com/)', 'Währung', 'Positionen Stücke', 'Transaktionskosten, Unrealisiert',
        'Kundenendbetrag ohne Transaktstionskosten', 'Preis (von https://finance.yahoo.com/)', 'Kurswert', 'Papiergewinne', 'Preis, T-30', '%%, T vs T-30',
        'Preis, T-90', '%%, T vs T-90', 'Preis, T-365', '%%, T vs T-365', 'Kundenendbetrag ohne Transaktstionskosten, abs'];

    let keysNewSheet = Array.from(resultsForNewSheet.keys());

    // check if list with name Ergebnisse (script) is already exist
    let name = "Erg.(script)";
    let nowDate = new Date();
    nowDate.setHours(nowDate.getHours() + 2);
    if (ss.getSheetByName(name) !== null) {
        name += " " + nowDate.toISOString().slice(0, 16);
    }


    let newSheet = ss.insertSheet(name);
    newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const writer = new ColumnWriter(newSheet, 1, 2);

    for (let i = 0; i < keysNewSheet.length; i++) {
        let ISINColumn = writer.write(keysNewSheet[i].split(" ")[0]);
        let columnWithSymbol = writer.setFormula("=SEARCH_SYMBOL_YAHOO(" + columnToLetter(ISINColumn) + writer.row + ")");
        writer.write(keysNewSheet[i].split(" ")[1]);
        let posColumn = writer.write(currentBalace.get(keysNewSheet[i]));
        let clearBufCost = 0;
        let bufCost = 0;
        for (let j = 0; j < ISINrecords.get(keysNewSheet[i]).length; j++) {
            clearBufCost += ISINrecords.get(keysNewSheet[i])[j].get("clearKurs") * ISINrecords.get(keysNewSheet[i])[j].get("quantity");
            bufCost += ISINrecords.get(keysNewSheet[i])[j].get("kurs") * ISINrecords.get(keysNewSheet[i])[j].get("quantity");
        }
        writer.write(clearBufCost - bufCost);
        let priceRealColumn = writer.write(-clearBufCost);
        let kursColumn = writer.setFormula("=GET_PRICE_REAL_TIME_YAHOO(" + columnToLetter(columnWithSymbol) + writer.row + ")");
        let kursTotalColumn = writer.setFormula("=" + columnToLetter(kursColumn) + writer.row + "*" + columnToLetter(posColumn) + writer.row);
        writer.setFormula("=" + columnToLetter(kursTotalColumn) + writer.row + "+" + columnToLetter(priceRealColumn) + writer.row);
        let T_30Column = writer.setFormula("=GET_PRICE_30_DAYS_YAHOO(" + columnToLetter(columnWithSymbol) + writer.row + ")");
        writer.setFormula("=IF(" + columnToLetter(T_30Column) + writer.row + "=0" + ",," + columnToLetter(kursColumn) + writer.row + "/" + columnToLetter(T_30Column) + writer.row + "-1)");
        let T_90Column = writer.setFormula("=GET_PRICE_90_DAYS_YAHOO(" + columnToLetter(columnWithSymbol) + writer.row + ")");
        writer.setFormula("=IF(" + columnToLetter(T_90Column) + writer.row + "=0" + ",," + columnToLetter(kursColumn) + writer.row + "/" + columnToLetter(T_90Column) + writer.row + "-1)");
        let T_365Column = writer.setFormula("=GET_PRICE_365_DAYS_YAHOO(" + columnToLetter(columnWithSymbol) + writer.row + ")");
        writer.setFormula("=IF(" + columnToLetter(T_365Column) + writer.row + "=0" + ",," + columnToLetter(kursColumn) + writer.row + "/" + columnToLetter(T_365Column) + writer.row + "-1)");
        writer.setFormula("=abs(" + columnToLetter(priceRealColumn) + writer.row + ")");

        writer.nextRow();
    }

    // two decimal places for all cells
    newSheet.getRange(2, 1, newSheet.getLastRow() - 1, headers.length).setNumberFormat("#,##0.00");

    newSheet.getRange(2, 11, newSheet.getLastRow() - 1, 1).setNumberFormat("#%");
    newSheet.getRange(2, 13, newSheet.getLastRow() - 1, 1).setNumberFormat("#%");
    newSheet.getRange(2, 15, newSheet.getLastRow() - 1, 1).setNumberFormat("#%");

    rules = [
        SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setBackground('#F4D03F'),
        SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground('#F4D03F'),
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('-').setBackground('#F4D03F'),
        SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#58D68D'),
        SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground('#EC7063'),
    ]

    // apply formatting
    allRules = newSheet.getConditionalFormatRules();
    for (let i = 0; i < rules.length; i++) {
        allRules.push(rules[i].setRanges([newSheet.getRange(2, 1, newSheet.getLastRow() - 1, newSheet.getLastColumn())]).build());
    }
    newSheet.setConditionalFormatRules(allRules);

    // resize all columns to fit content
    newSheet.autoResizeColumns(1, headers.length);
}

let processDeal = (records, balance, lastDeal) => {
    let result = 0;
    let newBalance = balance;

    //Logger.log("lastDeal: %s", lastDeal.get("kurs"));

    for (let i = 0; i < records.length; i++) {
        if (Math.abs(records[i].get("quantity")) >= Math.abs(lastDeal.get("quantity"))) {
            //Logger.log("%s %s", lastDeal.get("kurs"), records[i].get("kurs"))

            result += (lastDeal.get("kurs") - records[i].get("kurs")) * Math.abs(lastDeal.get("quantity"));
            newBalance += lastDeal.get("quantity");
            //Logger.log("newBalance: %s ; result: %s ; recordQuant: %s", newBalance, result, records[i].get("quantity"));
            records[i].set("quantity", records[i].get("quantity") + lastDeal.get("quantity"));

            return [records, newBalance, result];
        }
        else {
            result += (lastDeal.get("kurs") - records[i].get("kurs")) * Math.abs(records[i].get("quantity"));
            newBalance -= records[i].get("quantity");
            lastDeal.set("quantity", lastDeal.get("quantity") + records[i].get("quantity"));
            //Logger.log("newBalance: %s ; result: %s ; recordQuant: %s", newBalance, result, records[i].get("quantity"));
            records.splice(i, 1);
            i--;
        }
        //Logger.log("lastDeal: %s", lastDeal.get("quantity"));

        if (lastDeal.get("quantity") === 0) {
            //Logger.log("lastDeal == 0 -- newBalance: %s ; result: %s ; recordQuant: %s", newBalance, result, records[i].get("quantity"));

            return [records, newBalance, result];
        }
    }

    // write the rest of the lastDeal in records
    records.push(lastDeal);
    newBalance -= lastDeal.get("quantity");

    //Logger.log(newBalance);
    return [records, newBalance, result];
}

let searchForSymbolYahoo = (inData) => {
    // Search for a symbol on Yahoo Finance.
    //
    // Args:
    //     inData (string): The share (ex. ISIN) to search for.
    //
    // Returns:
    //     string: The symbol for the share, or undefined if not found.
    //
    let res = undefined;
    let data;
    while (true) {

        let url = 'https://query1.finance.yahoo.com/v1/finance/search?q=' + inData + '&quotesCount=1&newsCount=0&listsCount=0&quotesQueryId=tss_match_phrase_query';

        let headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.109 Safari/537.36',
        };

        let resp = UrlFetchApp.fetch(url, {'headers': headers});
        data = JSON.parse(resp.getContentText());

        if ('quotes' in data && data['quotes'].length > 0) {
            Logger.log(res);
            if (res === data['quotes'][0]['symbol']){
                return res;
            }
            res = data['quotes'][0]['symbol'];
            inData = data['quotes'][0]['shortname']
        } else {
            return res;
        }
    }
}

/**
 * Returns symbol from finance.yahoo.com
 * @param {string} data ISIN, notations or something else.
 * @return symbol from finance.yahoo.com
 * @customfunction
*/
function SEARCH_SYMBOL_YAHOO(data) {
    return searchForSymbolYahoo(data);
}

// Return the price of a stock on a given date.
// If the stock price is not available, return undefined.
let getPriceFromYahooHistoric = (symbol, date) => {
    if (symbol === "#ERROR!") return undefined;
    if (symbol === "#NUM!") return undefined;

    // Construct the URL for the Yahoo Finance.
    let url = "https://query1.finance.yahoo.com/v7/finance/download/" + symbol + "?period1=0&period2=9999999999&interval=1d&events=history&includeAdjustedClose=true";
    // Fetch the CSV data from the API.
    let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    // Parse the CSV data.
    let csvData = Utilities.parseCsv(response.getContentText());

    // Convert the date to the format used by the API.
    date = date.toISOString().slice(0, 10);

    // Search for the date in the CSV data.
    for (let i = 0; i < csvData.length; i++) {
        // Return the stock price if the date matches.
        if (csvData[i][0] === date) {
            return Number(csvData[i][4]);
        }
    }

    // Return undefined if the date was not found.
    return undefined;
}

/**
 * Returns price by date from finance.yahoo.com
 * @param {string} symbol symbol from finance.yahoo.com
 * @param {string} date date in format YYYY-MM-DD.
 * @return price by date from finance.yahoo.com
 * @customfunction
*/
function GET_PRICE_BY_DATE_YAHOO(symbol, date) {
    return getPriceFromYahooHistoric(symbol, Date(date));
}

// year must be a number
let getPriceFromYahooLastInYear = (symbol, year) => {
    if (symbol === "#ERROR!") return undefined;
    if (symbol === "#NUM!") return undefined;

    // Construct the URL for the Yahoo Finance.
    let url = "https://query1.finance.yahoo.com/v7/finance/download/" + symbol + "?period1=0&period2=9999999999&interval=1d&events=history&includeAdjustedClose=true";

    // Fetch the CSV data from the API.
    let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    // Parse the CSV data.
    let csvData = Utilities.parseCsv(response.getContentText());

    let flagYearFound = false;

    // Search for the date in the CSV data.
    for (let i = 0; i < csvData.length; i++) {
        // Return the stock price if the date matches.
        if (csvData[i][0].slice(0, 4) === year) {
            flagYearFound = true;
        } else if (csvData[i][0].slice(0, 4) !== year && flagYearFound) {
            return Number(csvData[i - 1][4]);
        }
    }

    // Return undefined if the date was not found.
    return undefined;
}

/**
 * Returns last price in year from finance.yahoo.com
 * @param {string} symbol ISIN, notations or something else.
 * @param {string} year year
 * @return last price in year from finance.yahoo.com
 * @customfunction
*/
function GET_PRICE_LAST_IN_YEAR_YAHOO(symbol, year) {
    return getPriceFromYahooLastInYear(symbol, year);
}

// year must be a number
let getPriceFromYahooFirstInYear = (symbol, year) => {
    if (symbol === "#ERROR!") return undefined;
    if (symbol === "#NUM!") return undefined;

    // Construct the URL for the Yahoo Finance.
    let url = "https://query1.finance.yahoo.com/v7/finance/download/" + symbol + "?period1=0&period2=9999999999&interval=1d&events=history&includeAdjustedClose=true";

    // Fetch the CSV data from the API.
    let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    // Parse the CSV data.
    let csvData = Utilities.parseCsv(response.getContentText());

    // Search for the date in the CSV data.
    for (let i = 0; i < csvData.length; i++) {
        // Return the stock price if the date matches.
        if (csvData[i][0].slice(0, 4) === year) {
            return Number(csvData[i][4]);
        }
    }

    // Return undefined if the date was not found.
    return undefined;
}

/**
 * Returns first price in year from finance.yahoo.com
 * @param {string} symbol symbol from finance.yahoo.com
 * @param {string} year year
 * @return number price in year from finance.yahoo.com
 * @customfunction
*/
function GET_PRICE_FIRST_IN_YEAR_YAHOO(symbol, year) {
    return getPriceFromYahooFirstInYear(symbol, year);
}

let getPriceFromYahooRealTime = (symbol) => {
    if (symbol === "#ERROR!") return undefined;
    if (symbol === "#NUM!") return undefined;

    let url = "https://finance.yahoo.com/quote/" + symbol + "?p=" + symbol;
    let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    // Get the HTML content of the page
    let contentText = response.getContentText();

    // Find the string that contains the price
    let startInd = contentText.indexOf(`<fin-streamer class="Fw(b) Fz(36px) Mb(-4px) D(ib)" data-symbol="`) + 66;
    startInd = contentText.indexOf(`value="`, startInd) + 7;
    let endInd = contentText.indexOf(`"`, startInd + 1);

    // Extract the price from the string and convert it to a number
    //Logger.log(startInd);
    //Logger.log(price);
    //Logger.log(contentText.substring(startInd, endInd));
    return Number(contentText.substring(startInd, endInd));
}

/**
 * Returns price from finance.yahoo.com
 * @param {string} symbol symbol from finance.yahoo.com
 * @return number from finance.yahoo.com
 * @customfunction
*/
function GET_PRICE_REAL_TIME_YAHOO(symbol) {
    return getPriceFromYahooRealTime(symbol);
}

let columnToLetter = (column) => {
    var temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}

let addDealRes = (sheet, column, row, bufResult, taxes, ISINtax, year) => {
    sheet.getRange(row, column + 1).setValue(bufResult);
    sheet.getRange(row, column).setFormula("=(" + columnToLetter(column + 1) + row + "-" + columnToLetter(column + 4) + row + ")");
    sheet.getRange(row, column + 2, 1, 5).setFormulas(
        [[
            "=(" + columnToLetter(column + 1) + row + "*" + ISINtax + ")",
            "=(" + columnToLetter(column + 1) + row + "-" + columnToLetter(column + 2) + row + ")",
            "=(" + columnToLetter(column + 5) + row + "+" + columnToLetter(column + 6) + row + ")",
            "=(" + columnToLetter(column + 3) + row + "*" + taxes.get(year).kapital + ")",
            "=(" + columnToLetter(column + 3) + row + "*" + taxes.get(year).solidar + ")"
        ]]
    );
}

let getPrice30Days = (symbol) => {
    return getPriceFromYahooHistoric(symbol, new Date(new Date().setDate((new Date()).getDate() - 30)));
}

/**
 * Returns price from finance.yahoo.com now() - 30days
 * @param {string} symbol symbol from finance.yahoo.com
 * @return number from finance.yahoo.com
 * @customfunction
 */
function GET_PRICE_30_DAYS_YAHOO(symbol) {
    return getPrice30Days(symbol);
}

let getPrice90Days = (symbol) => {
    return getPriceFromYahooHistoric(symbol, new Date(new Date().setDate((new Date()).getDate() - 90)));
}

/**
 * Returns price from finance.yahoo.com now() - 90days
 * @param {string} symbol symbol from finance.yahoo.com
 * @return number from finance.yahoo.com
 */
function GET_PRICE_90_DAYS_YAHOO(symbol) {
    return getPrice90Days(symbol);
}

let getPrice365Days = (symbol) => {
    return getPriceFromYahooHistoric(symbol, new Date(new Date().setDate((new Date()).getDate() - 365)));
}

/**
 * Returns price from finance.yahoo.com now() - 365days
 * @param {string} symbol symbol from finance.yahoo.com
 * @returns number from finance.yahoo.com
 */
function GET_PRICE_365_DAYS_YAHOO(symbol) {
    return getPrice365Days(symbol);
}

// let getDataFromBoerseStuttgart = (ISIN) => {
//     let url = "https://www.boerse-stuttgart.de/api/bsg-feature-navigation/Search/PostSearchInput" + ISIN;
//     let params = {
//         "searchInput": ISIN,
//         "language": "en",
//         "datasource": "5849b3c3-7bd3-4570-9fed-df92b0788426"
//     }
//
//     let response = UrlFetchApp.fetch(url, {
//         'method': 'post',
//         'contentType': 'application/json',
//         'payload': JSON.stringify(params)
//     });
//
//     return response.getContentText();
// }

// -------------------------------------------------------------------------------------------------------------------------------------------

function getOnvistaPrice(fundCode) {
    // API-URL für die Kurse von Fonds von onvista
    const url = "https://www.onvista.de/fonds/" + fundCode + "/kurs";

    // HTTP-Anfrage an die API senden
    const response = UrlFetchApp.fetch(url);

    // Antwort der API abfragen
    if (response.getResponseCode() === 200) {
        // Kurs des Fonds aus der Antwort der API extrahieren
        // Kurs des Fonds zurückgeben
        return response.getContentText().split(":")[1].trim();
    } else {
        // Fehlermeldung ausgeben
        const errorMessage = response.getResponseCode() + ": " + response.getContentText();
        throw new Error(errorMessage);
    }
}

// // Beispielanwendung
// const fundCode = "LU0348798264";
// const fundPrice = getOnvistaPrice(fundCode);

// Kurs des Fonds in einer Zelle in Google Sheets ausgeben
// SheetApp.getActive().getRange("A1").setValue(fundPrice);
