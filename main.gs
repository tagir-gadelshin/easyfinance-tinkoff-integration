function onlyUnique(value, index, self) {
    return self.indexOf(value) === index
}
function getCountCells_(array, reduceCallback) {
    return array.reduce(reduceCallback, { count: 0, total: 0 });
}
function createFinancialReport() {
    let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    let sheet = SpreadsheetApp.getActiveSheet()
    const reduceCallback = (acc, cv) => {
        const length = cv.length;
        const count = cv.filter(cell => cell !== '').length;
        acc.count += count;
        acc.total += length;
        return acc;
    }
    const countCells = getCountCells_(
        sheet.getRange(`A1:A${sheet.getLastRow()}`).getValues(),
        reduceCallback
    )
    Logger.log(countCells.count, "rows will be processed")
    const numRows = countCells.count
    const numColumns = 5
    let comment = sheet.getRange("List!D2:D").getValues()
    let cards = sheet.getRange("List!B2:B").getValues()
    let cardsUnique = sheet.getRange(2, 2, numRows-1, 1).getValues().join().split(",").filter(onlyUnique)
    Logger.log(cardsUnique)
    const dict = {
        'Яндекс.Драйв': 'Каршеринг',
        'Яндекс Такси': 'Такси',
        'Штрафы ГИБДД': 'Платные дороги, штрафы',
        'Супермаркеты': 'Питание (до расчётов)',
        'Фастфуд': 'Питание (до расчётов)',
        'Рестораны': 'Питание (до расчётов)',
        'Жкх': 'Коммунальные платежи',
        'Московский транспорт': 'Общественный транспорт',
        'Интернет': 'Интернет и ТВ',
        'МТС': 'Сотовый телефон',
        'Parkomat': 'Стоянка',
        'Ремонт': 'Ремонт недвижимости'
    }
    const cardDict = {
        '*1422': 'cred',
        '*6138': 'cred',
        '*3705': 'cred',
        '*6925': 'yandex',
        '*7074': 'debet',
        '*5986': 'debet',
        '*3443': 'schet',
        '*7000': 'debet',
        '*9891': 'debet',
        '*8711': 'mobile',
        '': 'empty'
    }
    Logger.log("Parsing categories")
    for (let i = 0; i < numRows-1; i++) {
        for (let key in dict) {
            if (comment[i][0].indexOf(key) > -1) {
                sheet.getRange(i + 2, 5).setValue(dict[key])
                Logger.log(comment[i][0], " is ",dict[key])
            }
        }
    }
    Logger.log("Starting card nums processing")
    for (let i = 0; i < numRows-1; i++) {
        if (cards[i][0] in cardDict) {
            Logger.log(cards[i][0]," is ",cardDict[cards[i][0]])
            sheet.getRange(i + 2, 2).setValue(cardDict[cards[i][0]])
        }
    }
    let cardTypesUnique = sheet.getRange(2, 2, numRows-1, 1).getValues().join().split(",").filter(onlyUnique)
    Logger.log("Card types are: ", cardTypesUnique)
    Logger.log("Creating sheets per card types")
    for (let cardN in cardTypesUnique) {
        let yourNewSheet = activeSpreadsheet.getSheetByName(cardTypesUnique[cardN])
        if (yourNewSheet != null) {
            activeSpreadsheet.deleteSheet(yourNewSheet)
        }
        yourNewSheet = activeSpreadsheet.insertSheet()
        yourNewSheet.setName(cardTypesUnique[cardN])
        Logger.log(cardTypesUnique[cardN]," list is created")
    }
    let cardTypes = sheet.getRange("List!B2:B").getValues()
    Logger.log("Starting moving rows")
    for (let i = 0; i < numRows-1; i++) {
        let targetSheet = activeSpreadsheet.getSheetByName(cardTypes[i][0])
        let source = sheet.getRange(i + 2, 1, 1, numColumns).getValues()
        let sourceLine = Array.from({length: numColumns})
        for (let i = 0; i < numColumns; i++) {
            sourceLine[i] = source[0][i]
        }
        Logger.log("Moving line to list",cardTypes[i][0], ". Line data:", sourceLine)
        targetSheet.appendRow(sourceLine)
    }
    Logger.log("Finished processing")
}