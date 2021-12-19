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
    Logger.log('%s rows will be processed', countCells.count)
    const numRows = countCells.count
    const numColumns = 5
    let comment = sheet.getRange("List!D2:D").getValues()
    let sum = sheet.getRange("List!C2:C").getValues()
    let cards = sheet.getRange("List!B2:B").getValues()
    let cardsUnique = sheet.getRange(2, 2, numRows-1, 1).getValues().join().split(",").filter(onlyUnique)
    Logger.log(cardsUnique)
    const dict = {
        'Яндекс.Драйв': 'Каршеринг',
        'Яндекс Такси': 'Такси',
        'ГИБДД': 'Платные дороги, штрафы',
        'Платная дорога':'Платные дороги, штрафы',
        'Транспорт Pvp': 'Платные дороги, штрафы',
        'Pvp No': 'Платные дороги, штрафы',
        'Супермаркеты': 'Питание (до расчётов)',
        'Фастфуд': 'Питание (до расчётов)',
        'Рестораны': 'Питание (до расчётов)',
        'Жкх': 'Коммунальные платежи',
        'Московский транспорт': 'Общественный транспорт',
        'Метрополитен': 'Общественный транспорт',
        'Интернет': 'Интернет и ТВ',
        'МТС': 'Сотовый телефон',
        'Parkomat': 'Стоянка',
        'Abakarov': 'Стоянка',
        'Ремонт': 'Ремонт недвижимости',
        'ремонт': 'Ремонт недвижимости',
        'Топливо': 'Топливо',
        'USD':'USD',
        'Комиссия за операцию':'Оплата услуг банка',
        'Мобильный +7':'Сотовый телефон',
        'Marks & Spencer':'Одежда',
        'Манго Страхование':'Страхование',
        'Спорттовары':'Спортивные товары'
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
        '*7560': 'debet',
        '*8711': 'mobile',
        '*7522': 'anna',
        '': 'empty'
    }
    Logger.log("Parsing categories")
    for (let i = 0; i < numRows-1; i++) {
        for (let key in dict) {
            if (comment[i][0].indexOf(key) > -1) {
                sheet.getRange(i + 2, 5).setValue(dict[key])
                Logger.log('%s is %s', comment[i][0], dict[key])
            } else if ((sum[i][0] > 0) && (comment[i][0].indexOf(key) < -1)) {
            sheet.getRange(i + 2, 5).setValue('Импорт (доход)')
            } else if ((sum[i][0] < 0) && (comment[i][0].indexOf(key) < -1)) {
            sheet.getRange(i + 2, 5).setValue('Импорт (расход)')
            } 
        }
    }
    Logger.log("Starting card nums processing")
    for (let i = 0; i < numRows-1; i++) {
        if ((cards[i][0] in cardDict) && (comment[i][0].indexOf('USD') > -1)) {
            Logger.log('%s with comment: %s is USD',cards[i][0],comment[i][0])
            sheet.getRange(i + 2, 2).setValue('USD')
        } else if (cards[i][0] in cardDict) {
            Logger.log('%s is %s',cards[i][0],cardDict[cards[i][0]])
            sheet.getRange(i + 2, 2).setValue(cardDict[cards[i][0]])
        } else {
          Logger.log('ERROR: %s is not recognized card in row %s. Please, add it in the dictionary first.',cards[i][0],i+2)
          return
        }
    }
    let cardTypesUnique = sheet.getRange(2, 2, numRows-1, 1).getValues().join().split(",").filter(onlyUnique)
    Logger.log('Card types are: %s', cardTypesUnique)
    Logger.log("Creating sheets per card types")
    for (let cardN in cardTypesUnique) {
        let yourNewSheet = activeSpreadsheet.getSheetByName(cardTypesUnique[cardN])
        if (yourNewSheet != null) {
            activeSpreadsheet.deleteSheet(yourNewSheet)
        }
        yourNewSheet = activeSpreadsheet.insertSheet()
        yourNewSheet.setName(cardTypesUnique[cardN])
        Logger.log('%s list is created', cardTypesUnique[cardN])
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
        Logger.log('Moving line to list %s. Line data: %s',cardTypes[i][0],sourceLine)
        targetSheet.appendRow(sourceLine)
    }
    Logger.log("Finished processing")
}
