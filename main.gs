function onlyUnique(value, index, self) {
    return self.indexOf(value) === index
}

function createFinancialReport() {
    let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    let sheet = SpreadsheetApp.getActiveSheet()
    const numRows = 100
    const numColumns = 10
    let comment = sheet.getRange("List!D2:D").getValues()
    let cards = sheet.getRange("List!B2:B").getValues()
    let cardsUnique = sheet.getRange(2, 2, numRows, 1).getValues().join().split(",").filter(onlyUnique)
    Logger.log(cardsUnique)
    const dict = {
        'Яндекс.Драйв':'Каршеринг',
        'Яндекс Такси':'Такси',
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
    for (let i = 0; i <= numRows; i++) {
        for (let key in dict) {
            if (comment[i][0].indexOf(key) > -1) {
                sheet.getRange(i + 2, 5).setValue(dict[key])
                Logger.log(comment[i][0])
            }
        }
    }
    Logger.log("Starting card nums processing")
    for (let i = 0; i <= numRows; i++) {
        if (cards[i][0] in cardDict) {
            Logger.log(cards[i][0])
            Logger.log(cardDict[cards[i][0]])
            sheet.getRange(i + 2, 2).setValue(cardDict[cards[i][0]])
        }
    }
    let cardTypesUnique = sheet.getRange(2, 2, numRows, 1).getValues().join().split(",").filter(onlyUnique)
    for (let cardN in cardTypesUnique) {
        let yourNewSheet = activeSpreadsheet.getSheetByName("Name of your new sheet")
        if (yourNewSheet != null) {
            activeSpreadsheet.deleteSheet(yourNewSheet)
        }
        yourNewSheet = activeSpreadsheet.insertSheet()
        if (cardTypesUnique[cardN] == '') {
            yourNewSheet.setName("empty")
            Logger.log(cardTypesUnique[cardN].toString())
        } else {
            yourNewSheet.setName(cardTypesUnique[cardN])
        }
    }
    let cardTypes = sheet.getRange("List!B2:B").getValues()
    Logger.log("Starting moving rows")
    for (let i = 0; i <= numColumns; i++) {
        Logger.log(cardTypes[i][0])
        let targetSheet = activeSpreadsheet.getSheetByName(cardTypes[i][0])
        Logger.log(targetSheet)
        let source = sheet.getRange(i + 2, 1, 1, numColumns).getValues()
        Logger.log(source)
        let sourceLine = Array.from({length: 10})
        for (let i = 0; i <= numRows; i++) {
            sourceLine[i] = source[0][i]
        }
        Logger.log(sourceLine)
        targetSheet.appendRow(sourceLine)
    }

}

