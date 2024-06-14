function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('###')
  .addItem('Статистика', 'showSTATInfo')
  .addToUi();
}


function getElementsByClassName(parent, className) {
  var elements = [];
  var children = parent.getAllContent();
  
  for (var i = 0; i < children.length; i++) {
    if (children[i].getAttribute('class') && children[i].getAttribute('class').getValue() == className) {
      elements.push(children[i]);
    }
    
    elements = elements.concat(getElementsByClassName(children[i], className));
  }
  
  return elements;
}

function findCell(searchValue, searchPool) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var column = sheet.getRange(searchPool); // Замените "A:A" на столбец, в котором вы хотите искать
  var values = column.getValues();

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == searchValue) {
      return i+1; // Возвращает номер строки, где найдено значение. Прибавляем 1, потому что индексация начинается с 0
    }
  }
}


function findValuesByIds(data, idsToFind, keyToExtract) {
  const results = {};
  idsToFind.forEach((id) => {
    const values = data[id];
    results[id] = values ? values[keyToExtract] : null;
  });
  return results;
}


function rangeValues(sheet,columnNumber){
  const last = sheet.getLastRow();
    //var lastRow = 17;
  var range = sheet.getRange(2, columnNumber, last-1, 1);
  var values = range.getValues();

  var nonEmptyValues =  values.filter(function(row) {
    return row[0] !== '' && row[0] !== undefined && row[0] !== null;
  });

  return nonEmptyValues.flat()
  
}

function showSTATInfo(){

const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const cell = sheet.getActiveCell();
const cellValue = cell.getValue();
const cellAddress = cell.getA1Notation();

if (Object.prototype.toString.call(cellValue) === '[object Date]') {

  const ui = SpreadsheetApp.getUi();
  ui.alert('Ячейка ' + cellAddress + ' содержит дату: ' + cellValue.toLocaleDateString());
  //const cellBelow = cell.offset(0, 0);
  const column = cell.getColumn();
  const dateString = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const apiKey = 'eyJhbGciOiJFUzI1NiIsImtpZCI6IjIwMjQwNTA2djEiLCJ0eXAiOiJKV1QifQ.eyJlbnQiOjEsImV4cCI6MTczMzg1ODg1NiwiaWQiOiI3ZDIwNjdiNi0yMDkyLTQ3NWMtOTMxNy0wY2Y1NWFiOGQ2YTYiLCJpaWQiOjIxNjYzNDM3LCJvaWQiOjI5NDUzOSwicyI6MTAwLCJzaWQiOiI3YWVlOWQyYi0yMmFlLTQ4ZTYtYjQxOS04ZGUwNTY2MjVkNDciLCJ0IjpmYWxzZSwidWlkIjoyMTY2MzQzN30.UMKmv72n2Jyyn1YSUKqXhHKJoJssYVRWEzLxJt70NGJBPAz1j8pn2ImDP4E2U6mkr1rLBv0QyVtI6Koe-q1b_Q';  // Замените на ваш ключ API

  getOrderStat(sheet, apiKey, dateString, column);
  getAutoCtrStats(sheet,apiKey,dateString,column);
  getSearchCtrStats(sheet,apiKey,dateString,column);
  getPrice(sheet,apiKey,dateString,column);
  
} else {

    const ui = SpreadsheetApp.getUi();
    ui.alert('Ошибка: Ячейка ' + cellAddress + ' не содержит дату.');

  }

}

function getPrice(sheet,apiKey,dateString,column){
  
  const apiUrl = 'https://statistics-api.wildberries.ru/api/v1/supplier/orders';
  const newApiUrl = apiUrl + '?dateFrom=' + dateString + '&flag=1';
  const options = {
    method: 'get',
    contentType: 'application/json',
    headers: {
      'Authorization': apiKey
    }
  };
  try {
    const response = UrlFetchApp.fetch(newApiUrl, options);
    const responseData = JSON.parse(response.getContentText());

    const targetNmId = rangeValues(sheet,1);
    const orderType = "Клиентский";

    const uniqueItems = {};
    responseData.forEach(item => {
        if (item.orderType === orderType && !uniqueItems[item.nmId]) {
        uniqueItems[item.nmId] = item;
      }
    });
    //Logger.log(uniqueItems);
    const key = 'finishedPrice'; 
    const results = findValuesByIds(uniqueItems, targetNmId, key);
    var k = 0;
    targetNmId.forEach((row, rowIndex) => {
      const nmID = row.toString(); 
      const matchingData = results[nmID];
      //Logger.log(matchingData);
      if (matchingData !== undefined) {
          // Если нашли совпадение, записываем значение в соответствующие ячейки
          sheet.getRange(rowIndex + 2 + 8* k + 8, column).setValue(Math.floor(matchingData*0.95));
          k+=1;
      }
    });

  } catch (error) {
    Logger.log('Ошибка запроса: ' + error);
  }

}


function getAutoCtrStats(sheet, apiKey, dateString, column){

  const idValues = rangeValues(sheet, 2);
  const apiUrl = 'https://advert-api.wb.ru/adv/v2/fullstats'
  //Logger.log(idValues);

  const requestBody = idValues.map(row => ({

      id: row,
      dates: [dateString, dateString]  // Преобразуем дату в формат YYYY-MM-DD

  }))
  const options = {

    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': apiKey
    },
    payload: JSON.stringify(requestBody)

  };
  //Logger.log(options);
  try {

    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseData = JSON.parse(response.getContentText());
    const extractedData = responseData.map(item => ({

      advertId: item.advertId.toString(),
      ctr: item.ctr

    }));
    Logger.log(extractedData);
    idValues.forEach((row, rowIndex) => {

    const advertId = row.toString(); // Получаем значение advertId из ячейки
    const matchingData = extractedData.find(item => item.advertId === advertId); // Ищем соответствующий объект в JSON
    if (matchingData) {

      var number = findCell(advertId,"B:B");
      sheet.getRange(number + 2, column).setValue(matchingData.ctr/100)

    }

    });
  } catch (error) {

    Logger.log('Ошибка запроса: ' + error);

  }
}

function getSearchCtrStats(sheet, apiKey, dateString, column){

  const idValues = rangeValues(sheet, 3);
  const apiUrl = 'https://advert-api.wb.ru/adv/v2/fullstats'
  Logger.log(idValues);

  const requestBody = idValues.map(row => ({

      id: row,
      dates: [dateString, dateString]  // Преобразуем дату в формат YYYY-MM-DD

  }))
  const options = {

    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': apiKey
    },
    payload: JSON.stringify(requestBody)

  };
  Logger.log(options);
  try {

    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseData = JSON.parse(response.getContentText());
    const extractedData = responseData.map(item => ({

      advertId: item.advertId.toString(),
      ctr: item.ctr

    }));
    Logger.log(extractedData);
    idValues.forEach((row, rowIndex) => {

    const advertId = row.toString(); // Получаем значение advertId из ячейки
    const matchingData = extractedData.find(item => item.advertId === advertId); // Ищем соответствующий объект в JSON
    if (matchingData) {

      // Если нашли совпадение, записываем значение sum в колонку D
      var number = findCell(advertId,"C:C");
      sheet.getRange(number + 1, column).setValue(matchingData.ctr/100)

    }

    });
  } catch (error) {

    Logger.log('Ошибка запроса: ' + error);

  }
}

function getOrderStat(sheet, apiKey, dateString, column){

  const idValues = rangeValues(sheet, 1);
  //Logger.log(idValues);
  const apiUrl = 'https://seller-analytics-api.wildberries.ru/api/v2/nm-report/detail';
  const dateFromCell = dateString;
  const beginDate = `${dateFromCell} 00:00:00`;
  const endDate = new Date(dateFromCell);
  endDate.setDate(endDate.getDate() + 1);
  const endDateString = `${endDate.getFullYear()}-${String(endDate.getMonth() + 1).padStart(2, '0')}-${String(endDate.getDate()).padStart(2, '0')} 00:00:00`;
  var pages=false;
  var nmIdOrdersSumRub;
  var k = 0;
  do {

    const page = k+1;
    const requestBody = {

      nmIDs: idValues,
      period: {

          begin: beginDate,
          end: endDateString

      },
      page: page

    };
    const options = {

      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': apiKey
      },
      payload: JSON.stringify(requestBody)

    };
    //Logger.log(options);
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseData = JSON.parse(response.getContentText());
    pages = responseData.data.isNextPage;
    nmIdOrdersSumRub = responseData.data.cards.map(card => ({

      nmID: card.nmID.toString(),
      ordersSumRub: card.statistics.selectedPeriod.ordersSumRub,
      openCardCount: card.statistics.selectedPeriod.openCardCount,
      ordersCount: card.statistics.selectedPeriod.ordersCount,
      addToCartPercent: card.statistics.selectedPeriod.conversions.addToCartPercent,
      cartToOrderPercent: card.statistics.selectedPeriod.conversions.cartToOrderPercent,
      //avgPriceRub: card.statistics.selectedPeriod.avgPriceRub

    }));

    //Logger.log(nmIdOrdersSumRub);
  }
  while (pages); 

  //Logger.log(idValues);
  var k = 0;

  idValues.forEach((row, rowIndex) => {

    const nmID = row.toString(); // Получаем значение advertId из ячейки
    const matchingData = nmIdOrdersSumRub.find(item => item.nmID === nmID); // Ищем соответствующий объект в JSON
    if (matchingData) {

      // Если нашли совпадение, записываем значение sum в колонку D
      sheet.getRange(rowIndex + 2 + 8*k, column).setValue(matchingData.openCardCount);
      sheet.getRange(rowIndex + 2 + 8*k + 3, column).setValue(matchingData.addToCartPercent/100);
      sheet.getRange(rowIndex + 2 + 8*k + 4, column).setValue(matchingData.cartToOrderPercent/100);
      sheet.getRange(rowIndex + 2 + 8*k + 5, column).setValue(matchingData.ordersCount);
      sheet.getRange(rowIndex + 2 + 8*k + 6, column).setValue(matchingData.ordersSumRub);

      if (matchingData.openCardCount != 0){

        sheet.getRange(rowIndex + 2 + 8*k + 7, column).setValue(matchingData.ordersCount/matchingData.openCardCount);

      } else sheet.getRange(rowIndex + 2 + 8*k + 7, column).setValue(0);
      //sheet.getRange(rowIndex + 2 + 7*k + 7, column).setValue(matchingData.avgPriceRub*0.82*0.95)
      k+=1;

    }

  });
  
}