//Создание меню 
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('###')
    .addItem('Статистика', 'showSTATInfo')
    .addItem('Текущий контроль', 'showADVERTinfo')
    .addToUi();
}
// Создает массив с артикулами, формирует запрос и отправляет его. Возвращает статистику по закзам и вносит данные в ячейки
function getORDERSTAT(column, dateString, apiKey, sheet){
  const idRange = sheet.getRange('A3:A33');  // Замените на ваш диапазон ID
  const idValues = idRange.getValues();
  const flatValues = idValues.flat();
  const apiUrl = 'https://seller-analytics-api.wildberries.ru/api/v2/nm-report/detail';
  const k = 0;
  const dateFromCell = dateString;
  const beginDate = `${dateFromCell} 00:00:00`;
  const endDate = new Date(dateFromCell);
  endDate.setDate(endDate.getDate() + 1);
  const endDateString = `${endDate.getFullYear()}-${String(endDate.getMonth() + 1).padStart(2, '0')}-${String(endDate.getDate()).padStart(2, '0')} 00:00:00`;
  const filteredIdValues = flatValues.filter(id => id !== 'НОВЫЕ ПОЗИЦИИ');
  Logger.log(filteredIdValues);
  var pages=false;
  var nmIdOrdersSumRub;
  do {
    const page = k+1;
    const requestBody = {
      nmIDs: filteredIdValues,
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
  Logger.log(options);
  const response = UrlFetchApp.fetch(apiUrl, options);
  const responseData = JSON.parse(response.getContentText());
  pages = responseData.data.isNextPage;
  nmIdOrdersSumRub = responseData.data.cards.map(card => ({
    nmID: card.nmID.toString(),
    ordersSumRub: card.statistics.selectedPeriod.ordersSumRub
  }));
  Logger.log(nmIdOrdersSumRub);
  }
  while (pages);
  idValues.forEach((row, rowIndex) => {
    const nmID = row[0].toString(); // Получаем значение advertId из ячейки
    const matchingData = nmIdOrdersSumRub.find(item => item.nmID === nmID); // Ищем соответствующий объект в JSON
    if (matchingData) {
      // Если нашли совпадение, записываем значение sum в колонку D
      sheet.getRange(rowIndex+3, column).setValue(matchingData.ordersSumRub); // rowIndex + 3, потому что диапазон начинается с C3
    }
    });
}
// Создает массив с id РК, формирует запрос и отправляет его. Возвращает статистику по РК и вносит данные в ячейки
function gETADVERSTAT(column, dateString, apiKey, sheet){
  const idRange = sheet.getRange('C3:C33');  // Замените на ваш диапазон ID
  const idValues = idRange.getValues();
  const apiUrl = 'https://advert-api.wb.ru/adv/v2/fullstats';
  const requestBody = idValues.map(row => ({
      id: row[0],
      dates: [dateString, dateString]  // Преобразуем дату в формат YYYY-MM-DD
    }))
    .filter(item => item.id !== ""); // Исключаем элементы с пустым id
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': apiKey
    },
    payload: JSON.stringify(requestBody)
  };
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseData = JSON.parse(response.getContentText());
    const extractedData = responseData.map(item => ({
      advertId: item.advertId.toString(),
      sum: item.sum
    }));
    idValues.forEach((row, rowIndex) => {
    const advertId = row[0].toString(); // Получаем значение advertId из ячейки
    const matchingData = extractedData.find(item => item.advertId === advertId); // Ищем соответствующий объект в JSON
    if (matchingData) {
      // Если нашли совпадение, записываем значение sum в колонку D
      sheet.getRange(rowIndex+3, column+1).setValue(matchingData.sum); // rowIndex + 3, потому что диапазон начинается с C3
    }
    });
  } catch (error) {
    Logger.log('Ошибка запроса: ' + error);
  }
}
// Отвечает за работу кнопки "Статистика". Проверяет является ли выбранная ячейка датой и запустает две функции
//gETADVERSTAT и getORDERSTAT
function showSTATInfo(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getActiveCell();
  const cellValue = cell.getValue();
  const cellAddress = cell.getA1Notation();
  if (Object.prototype.toString.call(cellValue) === '[object Date]') {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Ячейка ' + cellAddress + ' содержит дату: ' + cellValue.toLocaleDateString());
    const cellBelow = cell.offset(1, 0);
    const column = cellBelow.getColumn();
    const dateString = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const apiKey = 'eyJhbGciOiJFUzI1NiIsImtpZCI6IjIwMjQwNTA2djEiLCJ0eXAiOiJKV1QifQ.eyJlbnQiOjEsImV4cCI6MTczMjgzMjgwOSwiaWQiOiIyYmE1NDYwZC02ZGRlLTQzZmEtYjI5Yi04YWZlMzJmYzNlZDMiLCJpaWQiOjMzNzUxMDUzLCJvaWQiOjY5MjQ2LCJzIjo2OCwic2lkIjoiMjBiYzMxYjMtMzcwNi01YjUyLWFjZjktYjZhNGYzZTQ2N2NmIiwidCI6ZmFsc2UsInVpZCI6MzM3NTEwNTN9.6iXgzC6oZX7lVJThrcqHb8D2B-tfvNk061hpsXhglHYsc0m_EjGCcf8j_4ipnZsxVbNrzAMg9NDbEcHXyudTKQ';  // Замените на ваш ключ API
    getORDERSTAT(column, dateString, apiKey, sheet);
    gETADVERSTAT(column, dateString, apiKey, sheet);
    } else {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Ошибка: Ячейка ' + cellAddress + ' не содержит дату.');
    }

}
//Отвечает за работу кнопки "Текущий контроль". Вызывает функции getAdvertStatus и getAdvertbalance
function showADVERTinfo(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = 'eyJhbGciOiJFUzI1NiIsImtpZCI6IjIwMjQwNTA2djEiLCJ0eXAiOiJKV1QifQ.eyJlbnQiOjEsImV4cCI6MTczMjgzMjgwOSwiaWQiOiIyYmE1NDYwZC02ZGRlLTQzZmEtYjI5Yi04YWZlMzJmYzNlZDMiLCJpaWQiOjMzNzUxMDUzLCJvaWQiOjY5MjQ2LCJzIjo2OCwic2lkIjoiMjBiYzMxYjMtMzcwNi01YjUyLWFjZjktYjZhNGYzZTQ2N2NmIiwidCI6ZmFsc2UsInVpZCI6MzM3NTEwNTN9.6iXgzC6oZX7lVJThrcqHb8D2B-tfvNk061hpsXhglHYsc0m_EjGCcf8j_4ipnZsxVbNrzAMg9NDbEcHXyudTKQ';  // Замените на ваш ключ API
  const idRange = sheet.getRange('C3:C33');  // Замените на ваш диапазон ID
  const idValues = idRange.getValues();
  getAdvertStatus(idValues,sheet,apiKey);
  getAdvertbalance(sheet, idValues, apiKey);
}


//Возвращает состояние РК на данный момент по ее идентификатору
function getAdvertStatus(idValues,sheet,apiKey){
  const apiUrl = 'https://advert-api.wb.ru/adv/v1/promotion/adverts';
  const flatValues = idValues.flat();
  const filterVAlues = flatValues.filter(id => id !== "");
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': apiKey
    },
    payload: JSON.stringify(filterVAlues)
  };
  const response = UrlFetchApp.fetch(apiUrl, options);
  const responseData = JSON.parse(response.getContentText());
  const extractedData = responseData.map(item => ({
    advertId: item.advertId.toString(),
    status: item.status
  }));
  idValues.forEach((row, rowIndex) => {
    const advertId = row[0].toString(); // Получаем значение advertId из ячейки
    const matchingData = extractedData.find(item => item.advertId === advertId); // Ищем соответствующий объект в JSON
    if (matchingData) {
      // Если нашли совпадение, записываем значение sum в колонку D
      if (matchingData.status === 9){
        sheet.getRange(rowIndex+3, 5).setValue("А"); // rowIndex + 3, потому что диапазон начинается с C3
      }
      else {
        sheet.getRange(rowIndex+3, 5).setValue("П"); // rowIndex + 3, потому что диапазон начинается с C3
      }
    }
  });

}


//Возвращает баланс РК в данный момент по ее идентификатору.
function getAdvertbalance(sheet, idValues, apiKey){
  const apiUrl = 'https://advert-api.wb.ru/adv/v1/budget';
  const flatValues = idValues.flat();
  const filterVAlues = flatValues.filter(id => id !== "");
  let idBalance = [];
  function addObject(array, attr1, attr2) {
    const newObject = {
      id: attr1.toString(),
      total: attr2.toString()
    };
    array.push(newObject);
  }
  filterVAlues.forEach(id => {
    const apiURLnew = apiUrl +'?id='+id;
    const options = {
      method: 'get',
      contentType: 'application/json',
      headers: {
        'Authorization': apiKey
      }
    };
    const response = UrlFetchApp.fetch(apiURLnew, options);
    const responseData = JSON.parse(response.getContentText());
    const total = responseData.total;
    addObject(idBalance,id, total);
  });
  idValues.forEach((row, rowIndex) => {
    // Проверяем, есть ли такой id в полученных данных
    const balanceid = row[0].toString();
    const balanceData = idBalance.find(item => item.id === balanceid);
    // Если данные найдены, записываем значение total в соответствующую ячейку
    if (balanceData) {
      const cell = sheet.getRange(rowIndex+3, 4); // индексация в Google Sheets начинается с 1
      cell.setValue(balanceData.total);
    }
  });

}
