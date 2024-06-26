function calculateWorkingHoursByWeek(workTimeTable) {
    let workingHoursByWeek = [];
    workingHoursByWeek.push("Working Hours");
    let week_num = 0;

    workTimeTable.forEach(week => {
        let minutesForWeek = 0;

        
        week.forEach(entry => {
            let workingHours = entry?.workingHours;

            // Проверяем, что в записи есть информация о времени работы
            if (workingHours && parseInt(workingHours)) {
              let timeRanges = getSeparatedTimeRanges(workingHours);

              // Проходимся по каждому временному интервалу и вычисляем общее время работы
              timeRanges.forEach(timeRange => {
                  let [start, end] = timeRange.split('-');
                  let [startHours, startMinutes] = parseTime(start);
                  let [endHours, endMinutes] = parseTime(end);
                  
                  // Вычисляем разницу между временем начала и временем окончания в минутах
                  let totalMinutes = (endHours * 60 + endMinutes) - (startHours * 60 + startMinutes);

                  Logger.log(minutesToTime(totalMinutes))
                  minutesForWeek += totalMinutes;
              });
            }
        });

        workingHoursByWeek.push(minutesToTime(minutesForWeek));
        Logger.log('week %s', week_num);
        Logger.log('week %s', week);
        Logger.log('totalWorkTimeByWeek[week] %s',workingHoursByWeek[week_num]);
        week_num += 1;
    });

    return workingHoursByWeek;
}


function calculatePayment(workingHoursByWeek, hourlyWage) {
    let payment = [];
    payment.push("Expected Payment");

    workingHoursByWeek.forEach(time => {
        let [hours, minutes] = time.split(':').map(Number);
        let totalMinutes = hours * 60 + minutes;
        let roundedMinutes = Math.floor(totalMinutes / 30) * 30; // Округляем вниз до ближайшего отрезка по 30 минут
        payment.push(roundedMinutes / 60 * hourlyWage);
    });

    return payment;
}

function parseTime(timeString) {
  // Ищем символ между цифрами, если есть
  let match = timeString.match(/(\d+)(\D)(\d+)/);
  if (match) {
      let hours = parseInt(match[1]);
      //let separator = match[2];
      let minutes = parseInt(match[3]) || 0;
      return [hours, minutes];
  }
  return[parseInt(timeString), 0];
}

function getSeparatedTimeRanges(workingHours) {
  // Если временные интервалы разделены точкой с запятой
  if (workingHours.includes(';')) {
      return workingHours.split(';');
  }
  // Если временные интервалы разделены запятой
  else if (workingHours.includes(',')) {
      return workingHours.split(',');
  }
  // Если временных интервалов нет, возвращаем массив с одним элементом - исходной строкой
  else {
      return [workingHours];
  }
}


function minutesToTime(minutes) {
    // Вычисляем часы и минуты
    let resultHours = Math.floor(minutes / 60);
    let resultMinutes = minutes % 60;

    // Форматируем результат
    return `${resultHours}:${resultMinutes.toString().padStart(2, '0')}`;
}
