function bob () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.insertSheet('My New Sheet');
}
/**
 * Evaluates results by FIFO.
 * @param {Array<string>} dates dates of deals.
 * @param {Array<string>} ISINs ISINs of deals.
 * @param {Array<string>} notations notations of deals.
 * @param {Array<string>} typesOfDeals types of deals.
 * @param {Array<number>} quantities quantities for each deal.
 * @param {Array<number>} kurses kurses for each deal.
 * @param {Array<number>} cost cost for each deal.
 * @param {Array<string>} currencies currency for each deal.
 * @param {Array<number>} yearsTaxes years for taxes.
 * @param {Array<number>} taxBundesbankzinen Bundesbankzinen for each year.
 * @param {Array<number>} taxBasiszins Basiszins for each year.
 * @param {Array<number>} tax tax for each year.
 * @return Columns and rows with realized results and new sheet with unrealized results.
 * @customfunction
 * @OnlyCurrentDoc
*/
function EVAL_FIFO(dates, ISINs, notations, typesOfDeals, quantities, kurses, cost, currencies, yearsTaxes, taxBundesbankzinen, taxBasiszins, tax) {
    // check if all input arguments have the same length
    if (!(dates.length == ISINs.length && ISINs.length == notations.length && notations.length == typesOfDeals.length && typesOfDeals.length == kurses.length && kurses.length == cost.length && cost.length == currencies.length)) {
        throw "input values must have the same length";
    }

    // cehck if all taxes have the same length
    if (!(yearsTaxes.length == taxBundesbankzinen.length && taxBundesbankzinen.length == taxBasiszins.length && taxBasiszins.length == tax.length)) {
        throw "taxes must have the same length";
    }

    // make Map of all taxes
    let taxes = new Map();

    for (let i = 0; i < yearsTaxes.length; i++) {
        taxes.set(Number(yearsTaxes[i]), { "bundesbankzinen": taxBundesbankzinen[i], "basiszins": taxBasiszins[i], "tax": tax[i] });
    }

    // getting all data from input aguments in an array of maps
    let data = [];
    let currentBalace = new Map();
    let resultsForNewSheet = new Map();
    let averageKurs = new Map();

    for (let i = 0; i < dates.length; i++) {
        let bufMap = new Map();

        let dateParts = dates[i][0].split(".");
        let bufDate = new Date(dateParts[2], dateParts[1] - 1, dateParts[0]);

        if (!(bufDate instanceof Date)) {
            throw "element in dates is not a date";
        }

        bufMap.set("date", bufDate);
        bufMap.set("ISIN", ISINs[i][0]);
        bufMap.set("notation", notations[i][0]);
        bufMap.set("typeOfDeal", typesOfDeals[i][0]);

        if (typesOfDeals[i][0] !== "Kauf" && typesOfDeals[i][0] !== "Verkauf") {
            throw "element in typesOfDeals is not a Kauf or Verkauf";
        }

        // if (!isNaN(Number(quantities[i][0]))) {
        //     throw "element in quantities is not a number";
        // }

        // if (!isNaN(Number(cost[i][0]))) {
        //     throw "element in cost is not a number";
        // }

        if (typesOfDeals[i][0] === "Verkauf" && quantities[i][0] > 0) {
            bufMap.set("quantity", -quantities[i][0]);
        } else {
            bufMap.set("quantity", quantities[i][0]);
        }

        if (typesOfDeals[i][0] === "Verkauf" && cost[i][0] < 0) {
            bufMap.set("cost", -cost[i][0]);
            bufMap.set("kurs", Math.abs(cost[i][0] / bufMap.get("quantity")));
        } else {
            bufMap.set("cost", cost[i][0]);
            bufMap.set("kurs", Math.abs(cost[i][0] / bufMap.get("quantity")));
        }

        bufMap.set("currency", currencies[i][0]);

        data.push(bufMap);

        let key = ISINs[i][0] + " " + currencies[i][0]
        currentBalace.set(key, 0);

        if (resultsForNewSheet.get(key) === undefined) {
            resultsForNewSheet.set(key, new Map())
        } else {
            resultsForNewSheet.get(key).set(bufDate.getFullYear(), 0);
        }

        resultsForNewSheet.get(key).set("notation", notations[i][0]);

        if (averageKurs.get(key) === undefined) {
            averageKurs.set(key, [kurses[i][0]]);
        } else {
            averageKurs.set(key, [...averageKurs.get(key), kurses[i][0]]);
        }
    }

    // sort elemnts by dateinTypesOfDeals
    data.sort((a, b) => a.get("date") - b.get("date"));

    let keysNewSheet = Array.from(resultsForNewSheet.keys());

    for (let i = 0; i < keysNewSheet.length; i++) {
        let bufMap = resultsForNewSheet.get(keysNewSheet[i]);
        bufMap.set("symbol", searchForSymbolYahoo(resultsForNewSheet.get(keysNewSheet[i]).get("notation")))
        bufMap.set("kurs", getPriceFromYahooRealTime(bufMap.get("symbol")));
        bufMap.set("quantity", 0);
        bufMap.set("kurswert", 0);
        bufMap.set("kundenendbetrag", 0);
        bufMap.set("papiergewinne", 0);
    }

    // create array of years from data
    let years = [];
    for (let i = 0; i < data.length; i++) {
        if (years.indexOf(data[i].get("date").getFullYear()) === -1) {
            years.push(data[i].get("date").getFullYear());
        }
    }

    let ISINrecords = new Map();
    let result = [["Nettogewinne", ""]];

    for (let i = 0; i < data.length; i++) {
        let nowKey = data[i].get("ISIN") + " " + data[i].get("currency");


        if (currentBalace.get(nowKey) === 0) {
            // first deal with this ISIN
            //Logger.log("first deal with this ISIN")
            currentBalace.set(nowKey, data[i].get("quantity"));
            ISINrecords.set(nowKey, [data[i]]);
            result.push(["", ""]);
        } else if (currentBalace.get(nowKey) > 0 && data[i].get("typeOfDeal") === "Kauf") {
            // just add to current balance, no result
            //Logger.log("just add to current balance, no result")
            currentBalace.set(nowKey, currentBalace.get(nowKey) + data[i].get("quantity"));
            ISINrecords.get(nowKey).push(data[i]);
            result.push(["", ""]);
        } else if (currentBalace.get(nowKey) > 0 && data[i].get("typeOfDeal") === "Verkauf") {
            // need to evaluate result
            //Logger.log("need to evaluate result Verkauf")
            let records = ISINrecords.get(nowKey);
            let lastDeal = data[i];
            let balance = currentBalace.get(nowKey);
            let [bufRecords, bufBalance, bufResult] = processDeal(records, balance, lastDeal);
            ISINrecords.set(nowKey, bufRecords);
            currentBalace.set(nowKey, bufBalance);
            result.push([bufResult, ""]);
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
            result.push([bufResult, ""]);
            resultsForNewSheet.get(nowKey).set(data[i].get("date").getFullYear(), resultsForNewSheet.get(nowKey).get(data[i].get("date").getFullYear()) + bufResult);
        } else if (currentBalace.get(nowKey) < 0) {
            // just add to current balance, no result
            //Logger.log("just add to current balance, no result")
            currentBalace.set(nowKey, currentBalace.get(nowKey) + data[i].get("quantity"));
            ISINrecords.get(nowKey).push(data[i]);
            result.push(["", ""]);
        }
    }

    // write header
    let headers = ['ISIN', 'Bezeichnung', 'Positionen Stücke', 'Kundenendbetrag', 'Kurs (von https://finance.yahoo.com/)', 'Kurswert', 'Papiergewinne', 'Währung'];

    for (let i = 0; i < years.length; i++) {
        headers.push("Realisierte Gewinne/Verluste, " + years[i]);
    }

    for (let i = 0; i < years.length; i++) {
        headers.push("Nicht realisierte tax, " + years[i])
    }

    result[0].push(...headers);

    for (let i = 0; i < keysNewSheet.length; i++) {
        result[i + 1].push(keysNewSheet[i].split(" ")[0]);
        result[i + 1].push(resultsForNewSheet.get(keysNewSheet[i]).get("notation"));
        result[i + 1].push(currentBalace.get(keysNewSheet[i]));
        let bufCost = 0;
        for (let j = 0; j < ISINrecords.get(keysNewSheet[i]).length; j++) {
            bufCost += ISINrecords.get(keysNewSheet[i])[j].get("kurs") * ISINrecords.get(keysNewSheet[i])[j].get("quantity");
        }
        result[i + 1].push(-1 * bufCost);
        result[i + 1].push(resultsForNewSheet.get(keysNewSheet[i]).get("kurs"));
        result[i + 1].push(currentBalace.get(keysNewSheet[i]) * resultsForNewSheet.get(keysNewSheet[i]).get("kurs"));
        result[i + 1].push(currentBalace.get(keysNewSheet[i]) * resultsForNewSheet.get(keysNewSheet[i]).get("kurs") + -1 * bufCost);
        result[i + 1].push(keysNewSheet[i].split(" ")[1]);
        for (let j = 0; j < years.length; j++) {
            if (resultsForNewSheet.get(keysNewSheet[i]).get(years[j]) !== undefined) {
                result[i + 1].push(resultsForNewSheet.get(keysNewSheet[i]).get(years[i]));
            } else {
                result[i + 1].push("-");
            }
        }

        // counting tax
        for (let j = 0; j < years.length; j++) {
            if (resultsForNewSheet.get(keysNewSheet[i]).get(years[j]) !== undefined) {
                let firstInYear = getPriceFromYahooFirstInYear(resultsForNewSheet.get(keysNewSheet[i]).get("symbol"), years[j]);
                let lastInYear = getPriceFromYahooLastInYear(resultsForNewSheet.get(keysNewSheet[i]).get("symbol"), years[j]);
                let avg = averageKurs.get(keysNewSheet[i]).reduce((a, b) => a + b, 0) / averageKurs.get(keysNewSheet[i]).length;
                if (lastInYear < avg) {
                    result[i + 1].push(0)
                } else {
                    let tax = currentBalace.get(keysNewSheet[i]) * firstInYear * taxes.get(years[j]).basiszins * taxes.get(years[j]).tax;
                    result[i + 1].push(tax);
                }
            } else {
                result[i + 1].push("-");
            }
        }
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.insertSheet('My New Sheet');

    return result;
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
            Logger.log("lastDeal == 0 -- newBalance: %s ; result: %s ; recordQuant: %s", newBalance, result, records[i].get("quantity"));
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
    let url = "https://finance.yahoo.com/lookup?s=" + inData.replaceAll(" ", "%20");
    let res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    let contentText = res.getContentText();

    startInd = contentText.indexOf(`<td class="data-col0 Ta(start) Pstart(6px) Pend(15px)"><a href="/quote/`) + 71;
    if (startInd === -1) {
        return undefined;
    }
    for (let i = startInd; i < contentText.length; i++) {
        if (contentText[i] === "?") {
            return contentText.substring(startInd, i);
        }
    }
    return undefined;
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
    return searchForSymbolYahoo(symbol, year);
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
    return searchForSymbolYahoo(symbol, year);
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
