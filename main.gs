const COLUMNS_NAMES = {
    "dates": ["Abrechnungstag"],
    "ISIN": ["ISIN"],
    "typesOfDeals": ["Transaktion", "Geschäftsart"],
    "quantities": ["Stücke/Nom.", "Anteile"],
    "cost": ["Abrechnungsbetrag in Fondswährung", "Kundenendbetrag EUR"],
    "currencies": ["Fondswährung", "Währung"],
    "kurses": ["Abrechnungspreis", "Kurs"]
}

const DEAL_TYPE = {
    "buy": [
        "Kauf aus Sparplan",
        "Kauf",
        "Ertrag Wiederanlage",
        "Tausch (Kauf)",
        "Fondsmerge steuerpflichtig (Zugang)",
        "Kauf",
        "Tausch Gesamt (Kauf)",
        "Fondsmerge steuerneutral (Zugang)"],
    "sell": [
        "Tausch (Verkauf)",
        "Verkauf",
        "Tausch (Verkauf)",
        "Fondsmerge steuerpflichtig (Abgang)",
        "Vorabpauschale Verkauf",
        "Storno Verkauf",
        "Tausch Gesamt (Verkauf)",
        "Fondsmerge steuerneutral (Abgang)"],
    "ignore": [
        "Steuererstattung",
        "Ertrag",
        "Ertrag Auszahlung",
        "Vorabpauschale Abrechnung Lastschrift",
        "Steuererstattung",
        "Steuerforderung",
        "Vorabpauschale Abrechnung",
        "Entgeltbelastung",
        "Delta-Korrektur Abgang",
        "Storno Ertrag ohne Wiederanlage"]
}

const TAXES_SHEET_NAME = "taxes";

const TAXES_COLUMNS_NAMES = {
    "year": ["jahr", "year"],
    "tax": ["steuer", "tax"],
    "basiszins": ["basiszins"],
    "solidar": ["solidaritätszuschlag"],
    "kapital": ["kapitalertragsteuer"],
    "bbzinsen": ["bundesbank zinsen"],
    "basiszinsAnteil": ["basiszinssatz anteil"],
    "basiszinssatz": ["basiszinssatz"]
}


const ISIN_TAXES_SHEET_NAME = "isin_taxes";

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
    let data = sheet.getDataRange().getValues();

    let columns = { "dates": undefined, "ISIN": undefined, "typesOfDeals": undefined, "quantities": undefined, "cost": undefined, "currencies": undefined, "kurses": undefined };

    for (let i = 0; i < data.length; i++) {
        if (COLUMNS_NAMES.dates.includes(data[0][i])) {
            columns.dates = i;
        } else if (COLUMNS_NAMES.ISIN.includes(data[0][i])) {
            columns.ISIN = i;
        } else if (COLUMNS_NAMES.typesOfDeals.includes(data[0][i])) {
            columns.typesOfDeals = i;
        } else if (COLUMNS_NAMES.quantities.includes(data[0][i])) {
            columns.quantities = i;
        } else if (COLUMNS_NAMES.cost.includes(data[0][i])) {
            columns.cost = i;
        } else if (COLUMNS_NAMES.currencies.includes(data[0][i])) {
            columns.currencies = i;
        } else if (COLUMNS_NAMES.kurses.includes(data[0][i])) {
            columns.kurses = i;
        }
    }

    // check all columns are found and return missing column in alert
    let dataToEval = [];
    let keys = Object.keys(columns);

    for (let i = 0; i < keys.length; i++) {
        if (columns[keys[i]] === undefined) {
            //SpreadsheetApp.getUi().alert("column " + keys[i] + " is missing");
            throw "column " + keys[i] + " is missing";
        }

        let range = sheet.getRange(2, columns[keys[i]] + 1, sheet.getLastRow() - 1, 1);
        let data = range.getValues().flat();
        dataToEval.push(data);
    }

    // check if all input arguments have the same length
    let length = dataToEval[0].length;
    for (let i = 1; i < dataToEval.length; i++) {
        if (dataToEval[i].length !== length) {
            //SpreadsheetApp.getUi().alert("input arguments have different length");
            throw "input arguments have different length";
        }
    }

    if (sheet.getRange(1, sheet.getLastColumn()).getValue() !== "Nettogewinne (generated by FIFO script)") {
        sheet.insertColumnAfter(sheet.getLastColumn());
        evalFifo(dataToEval[0], dataToEval[1], dataToEval[2], dataToEval[3], dataToEval[4], dataToEval[5], dataToEval[6], sheet.getLastColumn() + 1);
    }
    else {
        evalFifo(dataToEval[0], dataToEval[1], dataToEval[2], dataToEval[3], dataToEval[4], dataToEval[5], dataToEval[6], sheet.getLastColumn());
    }
}

let convertStrToDate = (str) => {
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

let evalFifo = (dates, ISINs, typesOfDeals, quantities, cost, currencies, kurses, columnToPasteRes) => {
    // getting all data from input aguments in an array of maps
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    let data = [];
    let currentBalace = new Map();
    let resultsForNewSheet = new Map();
    let averageKurs = new Map();
    let yearStockBalance = new Map();
    let ISINsUnique = new Set();

    for (let i = 0; i < dates.length; i++) {
        let bufMap = new Map();

        let bufDate = convertStrToDate(dates[i]);

        if (isNaN(bufDate)) {
            //SpreadsheetApp.getUi().alert("element in dates is not a date (row: " + (i + 1) + ")");
            throw "element in dates is not a date (row: " + (i + 2) + ")";
        }

        bufMap.set("date", bufDate);
        bufMap.set("ISIN", ISINs[i]);

        // figure out type of deal
        if (DEAL_TYPE.buy.includes(typesOfDeals[i])) {
            bufMap.set("typeOfDeal", "Kauf");
        } else if (DEAL_TYPE.sell.includes(typesOfDeals[i])) {
            bufMap.set("typeOfDeal", "Verkauf");
        } else if (!(DEAL_TYPE.ignore.includes(typesOfDeals[i]))) {
            //SpreadsheetApp.getUi().alert("element in typesOfDeals is not a valid type (row: " + (i + 1) + ")");
            throw "element in typesOfDeals is not a valid type (row: " + (i + 2) + ")";
        }

        if (isNaN(quantities[i])) {
            //SpreadsheetApp.getUi().alert("element in quantities is not a number (row: " + (i + 1) + ")");
            throw "element in quantities is not a number (row: " + (i + 2) + ")";
        }
        if (bufMap.get("typeOfDeal") === "Verkauf" && quantities[i] > 0) {
            bufMap.set("quantity", -quantities[i]);
        } else {
            bufMap.set("quantity", quantities[i]);
        }

        if (bufMap.get("typeOfDeal") === "Verkauf" && cost[i] < 0) {
            bufMap.set("cost", -cost[i]);
            bufMap.set("kurs", Math.abs(cost[i] / bufMap.get("quantity")));
        } else {
            bufMap.set("cost", cost[i]);
            bufMap.set("kurs", Math.abs(cost[i] / bufMap.get("quantity")));
        }

        bufMap.set("currency", currencies[i]);

        data.push(bufMap);

        let key = ISINs[i] + " " + currencies[i]
        currentBalace.set(key, 0);

        if (resultsForNewSheet.get(key) === undefined) {
            resultsForNewSheet.set(key, new Map())
        }
        resultsForNewSheet.get(key).set(bufDate.getFullYear(), 0);

        if (yearStockBalance.get(key) === undefined) {
            yearStockBalance.set(key, new Map());
        }
        yearStockBalance.get(key).set(bufDate.getFullYear(), 0);


        // resultsForNewSheet.get(key).set("notation", notations[i]);

        if (averageKurs.get(key) === undefined) {
            averageKurs.set(key, [kurses[i]]);
        } else {
            averageKurs.set(key, [...averageKurs.get(key), kurses[i]]);
        }

        ISINsUnique.add(ISINs[i]);
    }

    // sort elemnts by dateinTypesOfDeals
    data.sort((a, b) => a.get("date") - b.get("date"));

    // for (let i = 0; i < keysNewSheet.length; i++) {
    //     let bufMap = resultsForNewSheet.get(keysNewSheet[i]);
    //     bufMap.set("symbol", searchForSymbolYahoo(resultsForNewSheet.get(keysNewSheet[i]).get("notation")))
    //     bufMap.set("kurs", getPriceFromYahooRealTime(bufMap.get("symbol")));
    //     bufMap.set("quantity", 0);
    //     bufMap.set("kurswert", 0);
    //     bufMap.set("kundenendbetrag", 0);
    //     bufMap.set("papiergewinne", 0);
    // }

    // create array of years from data
    let years = [];
    for (let i = 0; i < data.length; i++) {
        if (years.indexOf(data[i].get("date").getFullYear()) === -1) {
            years.push(data[i].get("date").getFullYear());
        }
    }

    let taxes = new Map();
    // search for sheet name
    let sheets = ss.getSheets();
    let leftYears = [...years];
    let taxSheet = ss.getSheetByName(TAXES_SHEET_NAME)

    if (taxSheet === null) {
        throw "sheet " + TAXES_SHEET_NAME + " not found";
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
        throw "sheet " + TAXES_SHEET_NAME + " does not contain all columns, missing columns: " + leftColumns.join(', ');
    }

    let yearsRows = new Map();
    for (let i = 2; i < taxSheet.getLastRow() + 1; i++) {
        let year = Number(taxSheet.getRange(i, columnsAdresses.get("year")).getValue());

        if (isNaN(year)) {
            throw "sheet " + TAXES_SHEET_NAME + " contains not a number in row " + (i);
        }

        if (leftYears.indexOf(year) !== -1) {
            leftYears.splice(leftYears.indexOf(year), 1);
        } else {
            throw "sheet " + TAXES_SHEET_NAME + " contains year " + year + " twice, in row " + (i);
        }

        yearsRows.set(year, i);
    }

    if (leftYears.length > 0) {
        throw "sheet " + TAXES_SHEET_NAME + " does not contain all years, missing years: " + leftYears.join(', ');
    }

    // write all adresses in map
    for (let i = 0; i < years.length; i++) {
        taxes.set(years[i], {
            "tax": `'` + TAXES_SHEET_NAME + `'!` + columnToLetter(columnsAdresses.get("tax")) + yearsRows.get(years[i]),
            "basiszins": `'` + TAXES_SHEET_NAME + `'!` + columnToLetter(columnsAdresses.get("basiszins")) + yearsRows.get(years[i]),
            "solidar": `'` + TAXES_SHEET_NAME + `'!` + columnToLetter(columnsAdresses.get("solidar")) + yearsRows.get(years[i]),
            "kapital": `'` + TAXES_SHEET_NAME + `'!` + columnToLetter(columnsAdresses.get("kapital")) + yearsRows.get(years[i]),
            "bbzinsen": `'` + TAXES_SHEET_NAME + `'!` + columnToLetter(columnsAdresses.get("bbzinsen")) + yearsRows.get(years[i]),
            "basiszinsAnteil": `'` + TAXES_SHEET_NAME + `'!` + columnToLetter(columnsAdresses.get("basiszinsAnteil")) + yearsRows.get(years[i]),
            "basiszinssatz": `'` + TAXES_SHEET_NAME + `'!` + columnToLetter(columnsAdresses.get("basiszinssatz")) + yearsRows.get(years[i])
        });
    }

    // collect ISIN taxes
    let ISINtaxes = new Map();
    let leftISINs = [...ISINsUnique];
    let flagISINTaxSheetFound = false;

    for (let i = 0; i < sheets.length; i++) {
        if (sheets[i].getName() == ISIN_TAXES_SHEET_NAME) {
            flagISINTaxSheetFound = true;

            for (let j = 2; j <= sheets[i].getLastRow(); j++) {
                let ISIN = sheets[i].getRange(j, 1).getValue();
                let teilfreistellung = Number(sheets[i].getRange(j, 2).getValue());

                // check if teilfreistellung is a number
                if (isNaN(teilfreistellung)) {
                    throw "teilfreistellung in sheet isin_taxes is not a number (row: " + j + ")";
                }

                // check if ISIN is already was in sheet
                if (ISINtaxes.get(ISIN) !== undefined) {
                    throw "ISIN " + ISIN + " was already in sheet isin_taxes (second appearance in row: " + j + ")";
                }

                leftISINs.splice(leftISINs.indexOf(ISIN), 1);

                ISINtaxes.set(ISIN, teilfreistellung);
            }
        }
    }

    let ISINrecords = new Map();
    let result = [];

    // write result column 
    var sheet = ss.getActiveSheet();
    sheet.getRange(1, columnToPasteRes).setValue("Nettogewinne (generated by FIFO script)");
    for (let i = 1; i < result.length; i++) {
        sheet.getRange(i + 1, columnToPasteRes).setValue(result[i - 1]);
    }

    for (let i = 0; i < data.length; i++) {
        let nowKey = data[i].get("ISIN") + " " + data[i].get("currency");


        if (currentBalace.get(nowKey) === 0) {
            // first deal with this ISIN
            //Logger.log("first deal with this ISIN")
            currentBalace.set(nowKey, data[i].get("quantity"));
            ISINrecords.set(nowKey, [data[i]]);
            sheet.getRange(i + 2, columnToPasteRes).setValue("");
        } else if (currentBalace.get(nowKey) > 0 && data[i].get("typeOfDeal") === "Kauf") {
            // just add to current balance, no result
            //Logger.log("just add to current balance, no result")
            currentBalace.set(nowKey, currentBalace.get(nowKey) + data[i].get("quantity"));
            ISINrecords.get(nowKey).push(data[i]);
            sheet.getRange(i + 2, columnToPasteRes).setValue("");
        } else if (currentBalace.get(nowKey) > 0 && data[i].get("typeOfDeal") === "Verkauf") {
            // need to evaluate result
            //Logger.log("need to evaluate result Verkauf")
            let records = ISINrecords.get(nowKey);
            let lastDeal = data[i];
            let balance = currentBalace.get(nowKey);
            let [bufRecords, bufBalance, bufResult] = processDeal(records, balance, lastDeal);
            ISINrecords.set(nowKey, bufRecords);
            currentBalace.set(nowKey, bufBalance);
            sheet.getRange(i + 2, columnToPasteRes).setValue(bufResult);
            resultsForNewSheet.get(nowKey).set(data[i].get("date").getFullYear(), resultsForNewSheet.get(nowKey).get(data[i].get("date").getFullYear()) + bufResult);
        } else if (currentBalace.get(nowKey) < 0 && data[i].get("typeOfDeal") === "Kauf") {
            // need to evaluate result
            //Logger.log("need to evaluate result Kauf")
            let records = ISINrecords.get(nowKey);
            let lastDeal = data[i];
            let balance = currentBalace.get(nowKey);
            let [bufRecords, bufBalance, bufResult] = processDeal(records, balance, lastDeal);
            ISINrecords.set(nowKey, bufRecords);
            currentBalace.set(nowKey, bufBalance);
            sheet.getRange(i + 2, columnToPasteRes).setValue(bufResult);
            resultsForNewSheet.get(nowKey).set(data[i].get("date").getFullYear(), resultsForNewSheet.get(nowKey).get(data[i].get("date").getFullYear()) + bufResult);
        } else if (currentBalace.get(nowKey) < 0) {
            // just add to current balance, no result
            //Logger.log("just add to current balance, no result")
            currentBalace.set(nowKey, currentBalace.get(nowKey) + data[i].get("quantity"));
            ISINrecords.get(nowKey).push(data[i]);
            sheet.getRange(i + 2, columnToPasteRes).setValue("");
        }
        yearStockBalance.get(nowKey).set(data[i].get("date").getFullYear(), currentBalace.get(nowKey));
    }

    // write header
    let headers = ['ISIN', 'Währung', 'Positionen Stücke', 'Kundenendbetrag', 'Symbol (von https://finance.yahoo.com/)', 'Kurs (von https://finance.yahoo.com/)', 'Kurswert', 'Papiergewinne'];

    for (let i = 0; i < years.length; i++) {
        headers.push("Realisierte Gewinne/Verluste, " + years[i]);
    }

    for (let i = 0; i < years.length; i++) {
        headers.push("Price for 01 Jan " + years[i]);
        headers.push("Average kurs of buying in " + years[i]);
        headers.push("Price for 31 Dec " + years[i]);
        headers.push("Positionen Stücke " + years[i]);
        headers.push("Nicht realisierte tax, " + years[i])
    }

    let keysNewSheet = Array.from(resultsForNewSheet.keys());

    // check if list with name Ergebnisse (generated by script) is already exist
    let name = "Ergebnisse (generated by script)";
    if (ss.getSheetByName(name) !== null) {
        name += " " + new Date().toISOString();
    }


    let newSheet = ss.insertSheet(name);
    newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    for (let i = 0; i < keysNewSheet.length; i++) {
        let writeColumn = 1;
        let bufCost = 0;
        for (let j = 0; j < ISINrecords.get(keysNewSheet[i]).length; j++) {
            bufCost += ISINrecords.get(keysNewSheet[i])[j].get("kurs") * ISINrecords.get(keysNewSheet[i])[j].get("quantity");
        }

        newSheet.getRange(i + 2, writeColumn, 1, 4).setValues(
            [[
                keysNewSheet[i].split(" ")[0],
                keysNewSheet[i].split(" ")[1],
                currentBalace.get(keysNewSheet[i]),
                -1 * bufCost
            ]]
        )
        writeColumn += 4;

        newSheet.getRange(i + 2, writeColumn, 1, 1).setFormulaR1C1("=SEARCH_SYMBOL_YAHOO(R[0]C[-4])");
        let symbolColumn = writeColumn;
        writeColumn++;

        newSheet.getRange(i + 2, writeColumn, 1, 1).setFormulaR1C1("=GET_PRICE_REAL_TIME_YAHOO(R[0]C[-1])");
        writeColumn++;

        newSheet.getRange(i + 2, writeColumn, 1, 1).setFormulaR1C1("=RC[-4]*RC[-1]");
        writeColumn++;

        newSheet.getRange(i + 2, writeColumn, 1, 1).setFormulaR1C1("=RC[-1]-RC[-4]");
        writeColumn++;

        for (let j = 0; j < years.length; j++) {
            if (resultsForNewSheet.get(keysNewSheet[i]).get(years[j]) !== undefined) {
                newSheet.getRange(i + 2, writeColumn, 1, 1).setValue(resultsForNewSheet.get(keysNewSheet[i]).get(years[j]));
            } else {
                newSheet.getRange(i + 2, writeColumn, 1, 1).setValue("-");
            }
            writeColumn++;
        }

        for (let j = 0; j < years.length; j++) {
            newSheet.getRange(i + 2, writeColumn, 1, 1).setFormulaR1C1("=GET_PRICE_FIRST_IN_YEAR_YAHOO(R[0]C[" + (-(writeColumn - 5)) + "], " + years[j] + ")");
            writeColumn++;
            newSheet.getRange(i + 2, writeColumn, 1, 1).setValue(averageKurs.get(keysNewSheet[i]).reduce((a, b) => a + b, 0) / averageKurs.get(keysNewSheet[i]).length);
            writeColumn++;
            newSheet.getRange(i + 2, writeColumn, 1, 1).setFormulaR1C1("=GET_PRICE_LAST_IN_YEAR_YAHOO(R[0]C[" + (-(writeColumn - 5)) + "], " + years[j] + ")");
            writeColumn++;
            newSheet.getRange(i + 2, writeColumn, 1, 1).setValue(yearStockBalance.get(keysNewSheet[i]).get(years[j]));
            writeColumn++;
            newSheet.getRange(i + 2, writeColumn, 1, 1).setFormula("=IF(" + columnToLetter(writeColumn - 2) + (i + 2) + "<" + columnToLetter(writeColumn - 3) + (i + 2) + ", 0, " + columnToLetter(writeColumn - 1) + (i + 2) + "*" + columnToLetter(writeColumn - 4) + (i + 2) + "*" + taxes.get(years[j]).basiszins + "*" + taxes.get(years[j]).tax + ")");
            writeColumn++;

            // if (resultsForNewSheet.get(keysNewSheet[i]).get(years[j]) !== undefined) {
            //     let firstInYear = getPriceFromYahooFirstInYear(resultsForNewSheet.get(keysNewSheet[i]).get("symbol"), years[j]); //P0
            //     let lastInYear = getPriceFromYahooLastInYear(resultsForNewSheet.get(keysNewSheet[i]).get("symbol"), years[j]); //P2
            //     let avg = averageKurs.get(keysNewSheet[i]).reduce((a, b) => a + b, 0) / averageKurs.get(keysNewSheet[i]).length; //P1
            //     if (lastInYear < avg) {
            //         newSheet.getRange(i + 2, 9 + years.length + j, 1, 1).setValue(0)
            //     } else {
            //         let tax = currentBalace.get(keysNewSheet[i]) * firstInYear * taxes.get(years[j]).basiszins * taxes.get(years[j]).tax;
            //         newSheet.getRange(i + 2, 9 + years.length + j, 1, 1).setValue(tax);
            //     }
            // } else {
            //     newSheet.getRange(i + 2, 9 + years.length + j, 1, 1).setValue("-");
            // }
        }


        // newSheet.getRange(i + 2, 1, 1, headers.length).setValues([
        //     [
        //         keysNewSheet[i].split(" ")[0],
        //         keysNewSheet[i].split(" ")[1],
        //         currentBalace.get(keysNewSheet[i]),
        //         -1 * bufCost,
        //         ,
        //         currentBalace.get(keysNewSheet[i]) * resultsForNewSheet.get(keysNewSheet[i]).get("kurs"),
        //         currentBalace.get(keysNewSheet[i]) * resultsForNewSheet.get(keysNewSheet[i]).get("kurs") + -1 * averageKurs.get(keysNewSheet[i]).reduce((a, b) => a + b, 0) / averageKurs.get(keysNewSheet[i]).length,
        //     ]
        // ]);
    }

    // resize all columns to fit content
    newSheet.autoResizeColumns(1, headers.length);

    // apply formatting
    let rules = [
        SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setBackground('#F4D03F'),
        SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground('#F4D03F'),
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('-').setBackground('#F4D03F'),
        SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#58D68D'),
        SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground('#EC7063'),
    ]

    let allRules = newSheet.getConditionalFormatRules();
    for (let i = 0; i < rules.length; i++) {
        allRules.push(rules[i].setRanges([newSheet.getRange(2, 1, newSheet.getLastRow() - 1, newSheet.getLastColumn())]).build());
    }
    newSheet.setConditionalFormatRules(allRules);



    // result[0].push(...headers);

    // for (let i = 0; i < keysNewSheet.length; i++) {
    //     result[i + 1].push(keysNewSheet[i].split(" ")[0]);
    //     result[i + 1].push(resultsForNewSheet.get(keysNewSheet[i]).get("notation"));
    //     result[i + 1].push(currentBalace.get(keysNewSheet[i]));
    //     let bufCost = 0;
    //     for (let j = 0; j < ISINrecords.get(keysNewSheet[i]).length; j++) {
    //         bufCost += ISINrecords.get(keysNewSheet[i])[j].get("kurs") * ISINrecords.get(keysNewSheet[i])[j].get("quantity");
    //     }
    //     result[i + 1].push(-1 * bufCost);
    //     result[i + 1].push(resultsForNewSheet.get(keysNewSheet[i]).get("kurs"));
    //     result[i + 1].push(currentBalace.get(keysNewSheet[i]) * resultsForNewSheet.get(keysNewSheet[i]).get("kurs"));
    //     result[i + 1].push(currentBalace.get(keysNewSheet[i]) * resultsForNewSheet.get(keysNewSheet[i]).get("kurs") + -1 * bufCost);
    //     result[i + 1].push(keysNewSheet[i].split(" ")[1]);
    //     for (let j = 0; j < years.length; j++) {
    //         if (resultsForNewSheet.get(keysNewSheet[i]).get(years[j]) !== undefined) {
    //             result[i + 1].push(resultsForNewSheet.get(keysNewSheet[i]).get(years[i]));
    //         } else {
    //             result[i + 1].push("-");
    //         }
    //     }

    //     // counting tax
    //     for (let j = 0; j < years.length; j++) {
    //         if (resultsForNewSheet.get(keysNewSheet[i]).get(years[j]) !== undefined) {
    //             let firstInYear = getPriceFromYahooFirstInYear(resultsForNewSheet.get(keysNewSheet[i]).get("symbol"), years[j]);
    //             let lastInYear = getPriceFromYahooLastInYear(resultsForNewSheet.get(keysNewSheet[i]).get("symbol"), years[j]);
    //             let avg = averageKurs.get(keysNewSheet[i]).reduce((a, b) => a + b, 0) / averageKurs.get(keysNewSheet[i]).length;
    //             if (lastInYear < avg) {
    //                 result[i + 1].push(0)
    //             } else {
    //                 let tax = currentBalace.get(keysNewSheet[i]) * firstInYear * taxes.get(years[j]).basiszins * taxes.get(years[j]).tax;
    //                 result[i + 1].push(tax);
    //             }
    //         } else {
    //             result[i + 1].push("-");
    //         }
    //     }
    // }
}

let processDeal = (records, balance, lastDeal) => {
    let result = 0;
    let newBalance = balance;

    //Logger.log("lastDeal: %s", lastDeal.get("quantity"));

    for (let i = 0; i < records.length; i++) {
        if (Math.abs(records[i].get("quantity")) >= Math.abs(lastDeal.get("quantity"))) {
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

        if (lastDeal.get("quantity") == 0) {
            //Logger.log("lastDeal == 0 -- newBalance: %s ; result: %s ; recordQuant: %s", newBalance, result, records[i].get("quantity"));
            return [records, newBalance, result];
        }
    }

    // write the rest of the lastDeal in records
    records.push(lastDeal);
    newBalance -= lastDeal.get("quantity");

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
    let url = 'https://query1.finance.yahoo.com/v1/finance/search?q=' + inData + '&quotesCount=1&newsCount=0&listsCount=0&quotesQueryId=tss_match_phrase_query';

    let headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.109 Safari/537.36',
    };

    let resp = UrlFetchApp.fetch(url, { 'headers': headers });
    let data = JSON.parse(resp.getContentText());

    if ('quotes' in data && data['quotes'].length > 0) {
        return data['quotes'][0]['symbol'];
    } else {
        return undefined;
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
    // Construct the URL for the Yahoo Finance.

    let url = "https://query1.finance.yahoo.com/v7/finance/download/" + symbol + "?period1=0&period2=9999999999&interval=1d&events=history&includeAdjustedClose=true";
    // Fetch the CSV data from the API.
    let response = UrlFetchApp.fetch(url);

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
        if (csvData[i][0].slice(0, 4) == year) {
            flagYearFound = true;
        } else if (csvData[i][0].slice(0, 4) != year && flagYearFound) {
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
    // Construct the URL for the Yahoo Finance.
    let url = "https://query1.finance.yahoo.com/v7/finance/download/" + symbol + "?period1=0&period2=9999999999&interval=1d&events=history&includeAdjustedClose=true";

    // Fetch the CSV data from the API.
    let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    // Parse the CSV data.
    let csvData = Utilities.parseCsv(response.getContentText());

    // Search for the date in the CSV data.
    for (let i = 0; i < csvData.length; i++) {
        // Return the stock price if the date matches.
        if (csvData[i][0].slice(0, 4) == year) {
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
 * @return first price in year from finance.yahoo.com
 * @customfunction
*/
function GET_PRICE_FIRST_IN_YEAR_YAHOO(symbol, year) {
    return getPriceFromYahooFirstInYear(symbol, year);
}

let getPriceFromYahooRealTime = (symbol) => {
    let url = "https://finance.yahoo.com/quote/" + symbol + "?p=" + symbol;
    let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    // Get the HTML content of the page
    let contentText = response.getContentText();

    // Find the string that contains the price
    let startInd = contentText.indexOf(`<fin-streamer class="Fw(b) Fz(36px) Mb(-4px) D(ib)" data-symbol="`) + 66;
    startInd = contentText.indexOf(`value="`, startInd) + 7;
    let endInd = contentText.indexOf(`"`, startInd + 1);

    // Extract the price from the string and convert it to a number
    let price = Number(contentText.substring(startInd, endInd));
    Logger.log(startInd);
    Logger.log(price);
    Logger.log(contentText.substring(startInd, endInd));
    return price;
}

/**
 * Returns price from finance.yahoo.com
 * @param {string} symbol symbol from finance.yahoo.com
 * @return price from finance.yahoo.com
 * @customfunction
*/
function GET_PRICE_REAL_TIME_YAHOO(symbol) {
    return getPriceFromYahooRealTime(symbol);
}

columnToLetter = (column) => {
    var temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}
