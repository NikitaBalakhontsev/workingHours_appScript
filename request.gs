const GLOBAL_SPREADSHEET_ID = "private"
const LOCAL_SPREADSHEET_ID = "private"
const LOCAL_SHEET_NAME = "private"
const LOCAL_SPREADSHEET = SpreadsheetApp.openById(LOCAL_SPREADSHEET_ID);
const GLOBAL_SPREADSHEET = SpreadsheetApp.openById(GLOBAL_SPREADSHEET_ID);
const LOCAL_SHEET = LOCAL_SPREADSHEET.getSheetByName(LOCAL_SHEET_NAME);


function getShetsFromMain() {
  const globalSheets = GLOBAL_SPREADSHEET.getSheets();

  let sheetNames = [];
  for (var i = 0; i < globalSheets.length; i++) {
    sheetNames.push([globalSheets[i].getName()]);
  }
  return sheetNames;
}


function getEmployeeNames() {
  const requestSheetName = LOCAL_SHEET.getRange('A1').getValue();
  if( requestSheetName === ""){
    message = "В ячейке A1 не выбран месяц"
    console.error(message);
    showErrorDialog(message);
    return null;
  }


  
  const requestSheet = GLOBAL_SPREADSHEET.getSheetByName(requestSheetName);
  if( requestSheet.length === 0){
    message = "Не удалось получить список работников для месяца"
    console.error(message);
    showErrorDialog(message);
    return null;
  }
  
  const lastRow = requestSheet.getLastRow();
  let employeeNames = requestSheet.getRange("A2:A" + lastRow).getValues();
  

  Logger.log('employeeNames.length : %s', employeeNames.length);
  Logger.log(employeeNames);

  // Находим позицию первой пустой строки
  const emptyRowIndex = employeeNames.findIndex(function(row) {
    return row[0] === "";
  });
  // Удаляем все строки после первой пустой
  employeeNames.splice(emptyRowIndex);
  
  // Сортируем фио по алфавиту
  employeeNames.sort(function(a, b) {
    return a[0].localeCompare(b[0], 'ru');
  });

  return employeeNames;
}


function getData() {
  const requestSheetName = LOCAL_SHEET.getRange('A1').getValue();
  const requestEmployeeName = LOCAL_SHEET.getRange('B1').getValue();
  const requestSheet =  GLOBAL_SPREADSHEET.getSheetByName(requestSheetName);
  const lastRow = requestSheet.getLastRow();
  const employeeNames = requestSheet.getRange("A1:A" + lastRow).getValues();

  const index = employeeNames.findIndex(function(item) {
  return item[0] === requestEmployeeName;
  });

  let workingHours = getWorkingHours(requestSheet, index);
  let dates = getDates(requestSheet);
  let workHoursData = createWorkHoursData(dates, workingHours);

  Logger.log(requestEmployeeName);
  Logger.log(employeeNames);
  Logger.log(index);

  Logger.log(workingHours);
  Logger.log('workingHours.length : %s', workingHours.length);
  
  Logger.log(dates);
  Logger.log('dates.length : %s', dates.length);
  
  Logger.log(workHoursData);
  Logger.log('workHoursData : %s', workHoursData.length);
  return workHoursData;
}


function getWorkingHours(sheet, index) {
  const lastRow = sheet.getLastRow();
  const workingHours = sheet.getRange(index + 1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return workingHours.filter((value, index) => index % 2 === 0 && index > 1).map(value => value.toString());
}


function getDates(sheet) {
  let dates = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  dates = removeValuesBeforeDate(dates);
  dates = dates.filter((value, index) => index % 2 === 0);
  dates = updateDatesYear(dates);
  return dates;
}


function createWorkHoursData(dates, workingHours) {
  return dates.map((date, index) => ({
    date: date || null,
    workingHours: workingHours[index] || null
  }));
}




function main(){
  const data = getData();
  Logger.log('data %s', data);

  let calendar = createWorkTimeTable(data);
  Logger.log(calendar);
  Logger.log('calendarLength : %s', calendar.length);

  let startCell = LOCAL_SHEET.getRange('B6');

  //Если данные удалось получить
  if(calendar.length > 3) {
    printWorkTimeTable(calendar,startCell);
  }
  else return null;


  clearRangeAndAddBorders(startCell.getRow(), startCell.getColumn() + 5, 16, 3);

  let workingHoursByWeek = calculateWorkingHoursByWeek(calendar);
  Logger.log('workingHoursByWeek %s', workingHoursByWeek);
  printColumn(workingHoursByWeek, startCell, 0, 5, 3, "@");
  

  let hourlyWage =  parseInt(LOCAL_SHEET.getRange('C1').getValue());
  let payment = calculatePayment(workingHoursByWeek, hourlyWage);
  Logger.log('payment %s', payment);
  printColumn(payment, startCell, 0, 6, 3, "0");


}

function createWorkTimeTable(workHoursData) {
  const weeks = [];
  let currentWeek = new Array(5).fill(null);
  let needUpdate = false;
    
  workHoursData.forEach(entry => {
    const date = new Date(entry.date);
    const dayOfWeek = date.getDay() % 5; // Изменили порядок дней (0 - пятница, 1 - понедельник ..., 4 - четверг)
  
    if(currentWeek[dayOfWeek] !== null ) needUpdate = true;

    if(dayOfWeek == 4){
       currentWeek[dayOfWeek] = entry;
       needUpdate = true;
    }
    
    if(needUpdate){
      weeks.push(currentWeek);
      currentWeek = new Array(5).fill(null)
      needUpdate = false;
    }
    else{
      currentWeek[dayOfWeek] = entry;
      if(currentWeek[dayOfWeek - 1] === null) {
        if(weeks.length === 0)
        {
          currentWeek[dayOfWeek - 1] = {date:"previous month", workingHours:null};
        }
        else currentWeek[dayOfWeek - 1] = {date:"weekend", workingHours:null};
      }
    }
  });
  let isEmpty = currentWeek => !currentWeek.some(element => element !== null);
  if(!isEmpty) {
    weeks.push(currentWeek);
  }
  return weeks;
}

function printRow(values, startCell, offsetRow, offsetColumn) {
  let startRow = startCell.getRow() + offsetRow;
  let startColumn = startCell.getColumn() + offsetColumn;
  let length = parseInt(values.length);

  let range = LOCAL_SHEET.getRange(startRow , startColumn, 1, length);
  range.setValues([values]).setHorizontalAlignment('center');
}

function printColumn(values, startCell, offsetRow, offsetColumn, offsetBetweenRows, numberFormat) {
  let currentRow = startCell.getRow() + offsetRow;
  let column = startCell.getColumn() + offsetColumn;
  let cell = LOCAL_SHEET.getRange(currentRow, column);
  cell.setValue(values[0]).setHorizontalAlignment("center");
  values.shift();
  currentRow += offsetBetweenRows - 1;

  // Проходимся по каждому значению в массиве
  values.forEach(value => {
    let cell = LOCAL_SHEET.getRange(currentRow, column);
    cell.setNumberFormat(numberFormat);
    cell.setValue(value).setHorizontalAlignment('center');
    currentRow += offsetBetweenRows; // Добавляем отступ 
  });
}

function printListToCell(list, cell){
  const range = LOCAL_SHEET.getRange(cell);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(list).build();
  range.setDataValidation(rule);
}


function printWorkTimeTable(workTimeTable, startCell) {
  let startRow = startCell.getRow();
  let startColumn = startCell.getColumn();
  const weekdays = ['Fri', 'Mon', 'Tue', 'Wed', 'Thu'];
  
  clearRangeAndAddBorders(startRow, startColumn, 16, 5);
  printRow(weekdays, startCell, 0, 0);
  //let range = LOCAL_SHEET.getRange(startRow, startColumn, 1, 5);
  //range.setValues([weekdays]).setHorizontalAlignment('center');
  startRow++;

  let entry;
  let dateCell;
  let hoursCell;
  // Вывод дат и часов работы
  for (let i = 0; i < workTimeTable.length; i++) {
    let week = workTimeTable[i];
    for (let j = 0; j < week.length; j++) {
      entry = week[j];
      if (entry === null) continue;

      dateCell = LOCAL_SHEET.getRange(startRow, startColumn + j);
      hoursCell = LOCAL_SHEET.getRange(startRow + 1, startColumn + j);

      dateCell.setValue(entry.date);
      hoursCell.setValue(entry.workingHours);

      dateCell.setHorizontalAlignment("center");
      hoursCell.setHorizontalAlignment("center");
    }
    // Смещение для следующей строки
    startRow += 3;
  }
}


function printSheetNames() {
  const sheetNames = getShetsFromMain();
  if(sheetNames === null){ return null};

  printListToCell(sheetNames, "A1");

  Logger.log(sheetNames);
  Logger.log('sheetNames.length : %s', sheetNames.length);
}


function printEmoloyeeNames() {
  const currentEmployee = LOCAL_SHEET.getRange('B1').getValue();
  const employeeNames = getEmployeeNames();
  if(employeeNames === null){ return null};

  let clearCurrentEmployee = true;
  let employeeName;

  for(i = 0; i < employeeNames.length; i++) {
    employeeName = employeeNames[i].toString();
    if (currentEmployee === employeeName) {
      clearCurrentEmployee = false;
      break;
    }
  }

  if(clearCurrentEmployee) {
    LOCAL_SHEET.getRange('B1').clear();
  }
  
  printListToCell(employeeNames, "B1");
  Logger.log(employeeNames);
  Logger.log('employeeNames.length : %s', employeeNames.length);
}


function onChangeSheets(e) {
  var sheet = e.source.getActiveSheet();
  // Проверяем, что изменения произошли в определенной ячейке
  if (sheet.getName() === LOCAL_SHEET_NAME && sheet.getActiveRange().getA1Notation() === 'A1') {
    // Выполняем код при изменении значения ячейки A1
    printEmoloyeeNames();
    const employeeName = LOCAL_SHEET.getRange("B1").getValue();
    if(employeeName) {
      main();
    }

    Logger.log('User pick month, sheetName: %s', sheet.getName());
    Logger.log('ActiveRange: %s', sheet.getActiveRange().getA1Notation());
  }

  if (sheet.getName() === LOCAL_SHEET_NAME && sheet.getActiveRange().getA1Notation() === 'B1') {
    // Выполняем код при изменении значения ячейки A1
    //var startCell = sheet.getRange('B6');
    //clearRangeAndAddBorders(startCell.getRow(), startCell.getColumn(), 16, 5);
    main();
    Logger.log('User pick employee, sheetName: %s',sheet.getName());
    Logger.log('ActiveRange: %s', sheet.getActiveRange().getA1Notation());
  }
}




function clearRangeAndAddBorders(startRow, startColumn, numRows, numColumns) {
  // Очистка области
  var rangeToClear = LOCAL_SHEET.getRange(startRow, startColumn, numRows, numColumns);
  rangeToClear.clear();
  
  // Добавление границ к очищенной области
  var borderStyle = SpreadsheetApp.BorderStyle.SOLID;
  var borderColor = "#000000"; // Черный цвет
  rangeToClear.setBorder(true, true, true, true, false, false, borderColor, borderStyle);
}


function removeValuesBeforeDate(arr) {
    for (var i = 0; i < arr.length; i++) {
        if (arr[i] instanceof Date) {
            // Найден элемент типа Date, удаляем все элементы до него
            arr.splice(0, i);
            return arr;
        }
    }
    return arr;
}


function updateDatesYear(dates) {
    // Пройтись по каждой дате в массиве
    for (var i = 0; i < dates.length; i++) {
        // Заменить год на 2024
        dates[i].setFullYear(2024);

        // Обновить день недели
        var daysOfWeek = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
        var dayIndex = dates[i].getDay();
        var currentDayOfWeek = daysOfWeek[dayIndex];

        // Обновить день недели в строковом представлении
        var dateString = dates[i].toString();
        var updatedDateString = dateString.replace(dateString.substr(0, 3), currentDayOfWeek);
        dates[i] = new Date(updatedDateString);
    }

    return dates;
}


function showErrorDialog(message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert('Ошибка', message, ui.ButtonSet.OK);
}
