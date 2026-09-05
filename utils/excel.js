const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const checklist = require('../config/checklist.json');
const { getAllAudits } = require('./audit');

const REPORTS_DIR = path.join(__dirname, '..', 'reports');

// Создаем директорию для отчетов
if (!fs.existsSync(REPORTS_DIR)) {
  fs.mkdirSync(REPORTS_DIR);
}

/**
 * Сгенерировать полный отчет в Excel
 */
async function generateFullReport(auditData) {
  const workbook = new ExcelJS.Workbook();
  
  // 1. Титульный лист
  await createTitleSheet(workbook, auditData);
  
  // 2. Чек-лист с результатами
  await createChecklistSheet(workbook, auditData);
  
  // 3. Сводная таблица и графики
  await createSummarySheet(workbook, auditData);
  
  // 4. Детали по каждому рабочему месту
  await createWorkplaceDetailsSheet(workbook, auditData);
  
  // 5. Информация о фотофиксации
  await createPhotosSheet(workbook, auditData);
  
  // 6. Шаблон для ручного плана мероприятий
  await createActionPlanTemplate(workbook);
  
  // Сохраняем файл
  const reportName = `Отчет_5С_${auditData.workplace}_${new Date().toLocaleDateString('ru-RU').replace(/\./g, '-')}.xlsx`;
  const reportPath = path.join(REPORTS_DIR, reportName);
  
  await workbook.xlsx.writeFile(reportPath);
  
  return reportPath;
}

/**
 * Создать титульный лист
 */
async function createTitleSheet(workbook, auditData) {
  const worksheet = workbook.addWorksheet('Титульный лист');
  
  // Настройки страницы
  worksheet.properties.defaultRowHeight = 20;
  worksheet.columns = [
    { width: 20 },
    { width: 20 },
    { width: 20 },
    { width: 20 },
    { width: 20 },
    { width: 20 },
    { width: 20 },
    { width: 20 }
  ];
  
  // Заголовок
  const title = 'ОТЧЕТ ПО АУДИТУ 5С';
  const subtitle = 'Овощной участок - Рабочие места по резке овощей';
  
  worksheet.getCell('B2').value = title;
  worksheet.getCell('B2').font = { name: 'Arial', size: 20, bold: true };
  worksheet.getCell('B2').alignment = { horizontal: 'center' };
  
  worksheet.getCell('B3').value = subtitle;
  worksheet.getCell('B3').font = { name: 'Arial', size: 14, bold: true };
  worksheet.getCell('B3').alignment = { horizontal: 'center' };
  
  // Информация об аудите
  const auditInfo = [
    ['Объект аудита:', auditData.workplace || 'Овощной участок'],
    ['Дата проведения:', new Date().toLocaleDateString('ru-RU')],
    ['Аудитор:', auditData.auditor || ''],
    ['Количество рабочих мест:', auditData.workplaceCount || 5],
    ['Общий балл:', auditData.totalScore || 0]
  ];
  
  auditInfo.forEach((row, index) => {
    worksheet.getCell(`B${5 + index}`).value = row[0];
    worksheet.getCell(`B${5 + index}`).font = { name: 'Arial', size: 12, bold: true };
    worksheet.getCell(`C${5 + index}`).value = row[1];
    worksheet.getCell(`C${5 + index}`).font = { name: 'Arial', size: 12 };
  });
  
  // Оценка уровня
  const level = getLevelByScore(auditData.totalScore || 0);
  worksheet.getCell('B11').value = 'Уровень 5С:';
  worksheet.getCell('B11').font = { name: 'Arial', size: 12, bold: true };
  worksheet.getCell('C11').value = level;
  worksheet.getCell('C11').font = { name: 'Arial', size: 12, bold: true, color: getLevelColor(level) };
  
  // Логотип или декоративный элемент
  worksheet.getCell('F2').value = 'Методика МУ-52-2024';
  worksheet.getCell('F2').font = { name: 'Arial', size: 12, italic: true };
  worksheet.getCell('F2').alignment = { horizontal: 'right' };
  
  // Объединяем ячейки
  worksheet.mergeCells('B2:G2');
  worksheet.mergeCells('B3:G3');
  worksheet.mergeCells('C5:G5');
  worksheet.mergeCells('C6:G6');
  worksheet.mergeCells('C7:G7');
  worksheet.mergeCells('C8:G8');
  worksheet.mergeCells('C9:G9');
  worksheet.mergeCells('C11:G11');
}

/**
 * Создать лист с чек-листом
 */
async function createChecklistSheet(workbook, auditData) {
  const worksheet = workbook.addWorksheet('Чек-лист');
  
  // Заголовки
  worksheet.getCell('A1').value = 'ЧЕК-ЛИСТ АУДИТА 5С';
  worksheet.getCell('A1').font = { name: 'Arial', size: 16, bold: true };
  worksheet.getCell('A2').value = 'Оценка: 1 - выполнено, 0 - не выполнено';
  worksheet.getCell('A2').font = { name: 'Arial', size: 10, italic: true };
  
  // Заголовки таблицы
  const headers = ['№', 'Секция', 'Критерий', ...Array.from({length: auditData.workplaceCount || 5}, (_, i) => `Раб. место ${i+1}`), 'Макс. балл'];
  
  const headerRow = worksheet.addRow(headers);
  headerRow.eachCell(cell => {
    cell.font = { name: 'Arial', size: 10, bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  });
  
  // Заполняем данные
  let rowIndex = 3;
  let criteriaIndex = 1;
  
  for (const sectionId in checklist.sections) {
    const section = checklist.sections[sectionId];
    
    // Добавляем заголовок секции
    const sectionRow = worksheet.addRow([`${sectionId}`, section.name, '', '', '', '', '', '', '']);
    sectionRow.getCell(1).font = { bold: true };
    sectionRow.getCell(2).font = { bold: true };
    sectionRow.getCell(2).colSpan = 8;
    rowIndex++;
    
    // Добавляем критерии
    for (const criteria of section.criteria) {
      const row = worksheet.addRow([
        criteriaIndex++,
        `  ${criteria.id}`,
        criteria.text,
        ...Array.from({length: auditData.workplaceCount || 5}, (_, i) => 
          auditData.results?.[sectionId]?.[criteria.id]?.[i] || ''
        ),
        criteria.weight
      ]);
      
      row.eachCell(cell => {
        cell.font = { name: 'Arial', size: 10 };
        cell.alignment = { vertical: 'top' };
        cell.border = { left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      });
      
      rowIndex++;
    }
    
    // Добавляем строку итогов по секции
    const sectionScore = auditData.sectionScores?.[sectionId]?.score || 0;
    const sectionMaxScore = section.maxScore;
    
    const totalRow = worksheet.addRow([
      `Итого ${sectionId}:`,
      '',
      '',
      ...Array.from({length: auditData.workplaceCount || 5}, () => ''),
      `${sectionScore}/${sectionMaxScore}`
    ]);
    
    totalRow.getCell(1).font = { bold: true };
    totalRow.getCell(totalRow.cellCount).font = { bold: true };
    totalRow.eachCell(cell => {
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    });
    
    rowIndex++;
  }
  
  // Итоговый балл
  const finalRow = worksheet.addRow([
    'ОБЩИЙ БАЛЛ:',
    '',
    '',
    ...Array.from({length: auditData.workplaceCount || 5}, () => ''),
    `${auditData.totalScore || 0}/25`
  ]);
  
  finalRow.getCell(1).font = { bold: true, size: 12 };
  finalRow.getCell(finalRow.cellCount).font = { bold: true, size: 12 };
  finalRow.eachCell(cell => {
    cell.border = { top: { style: 'thick' }, left: { style: 'thin' }, bottom: { style: 'thick' }, right: { style: 'thin' } };
  });
  
  // Настраиваем ширину колонок
  worksheet.getColumn(1).width = 5;
  worksheet.getColumn(2).width = 10;
  worksheet.getColumn(3).width = 60;
  for (let i = 4; i <= (auditData.workplaceCount || 5) + 3; i++) {
    worksheet.getColumn(i).width = 12;
  }
  worksheet.getColumn((auditData.workplaceCount || 5) + 4).width = 12;
}

/**
 * Создать сводный лист с графиками
 */
async function createSummarySheet(workbook, auditData) {
  const worksheet = workbook.addWorksheet('Сводка');
  
  // Заголовок
  worksheet.getCell('A1').value = 'СВОДНАЯ ИНФОРМАЦИЯ ПО АУДИТУ 5С';
  worksheet.getCell('A1').font = { name: 'Arial', size: 16, bold: true };
  
  // Таблица по секциям
  const sectionHeaders = ['Секция', 'Баллы', 'Максимум', 'Процент', 'Уровень'];
  const sectionRow = worksheet.addRow(sectionHeaders);
  sectionRow.eachCell(cell => {
    cell.font = { bold: true };
    cell.alignment = { horizontal: 'center' };
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  });
  
  let sectionRowIndex = 3;
  for (const sectionId in checklist.sections) {
    const sectionScore = auditData.sectionScores?.[sectionId];
    const level = getLevelByScore(sectionScore?.percentage || 0);
    
    const row = worksheet.addRow([
      checklist.sections[sectionId].name,
      sectionScore?.score || 0,
      sectionScore?.maxScore || checklist.sections[sectionId].maxScore,
      `${sectionScore?.percentage || 0}%`,
      level
    ]);
    
    row.eachCell(cell => {
      cell.alignment = { horizontal: 'center' };
      cell.border = { left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    });
    
    // Цвет в зависимости от уровня
    if (sectionRowIndex <= 7) {
      row.getCell(5).font = { color: getLevelColor(level) };
    }
    
    sectionRowIndex++;
  }
  
  // Итог
  const totalRow = worksheet.addRow([
    'ОБЩИЙ БАЛЛ',
    auditData.totalScore || 0,
    25,
    `${Math.round(((auditData.totalScore || 0) / 25) * 100)}%`,
    getLevelByScore(auditData.totalScore || 0)
  ]);
  
  totalRow.eachCell(cell => {
    cell.font = { bold: true };
    cell.alignment = { horizontal: 'center' };
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thick' }, right: { style: 'thin' } };
  });
  
  // Настраиваем ширину колонок
  worksheet.getColumn(1).width = 25;
  worksheet.getColumn(2).width = 10;
  worksheet.getColumn(3).width = 12;
  worksheet.getColumn(4).width = 12;
  worksheet.getColumn(5).width = 15;
  
  // Добавляем графики (временно отключено - ExcelJS не поддерживает графики в этой версии)
  // await createCharts(workbook, auditData);
  
  // Добавляем простую таблицу с данными для графиков
  await createChartDataSheet(workbook, auditData);
}

/**
 * Создать лист с деталями по рабочим местам
 */
async function createWorkplaceDetailsSheet(workbook, auditData) {
  const worksheet = workbook.addWorksheet('Детали по местам');
  
  // Заголовок
  worksheet.getCell('A1').value = 'ДЕТАЛЬНЫЙ АНАЛИЗ ПО РАБОЧИМ МЕСТАМ';
  worksheet.getCell('A1').font = { name: 'Arial', size: 16, bold: true };
  
  const workplaceCount = auditData.workplaceCount || 5;
  
  // Таблица
  const headers = ['Рабочее место', ...Object.keys(checklist.sections).map(s => checklist.sections[s].name), 'Итого'];
  const headerRow = worksheet.addRow(headers);
  headerRow.eachCell(cell => {
    cell.font = { bold: true };
    cell.alignment = { horizontal: 'center' };
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  });
  
  // Данные по рабочим местам
  for (let i = 0; i < workplaceCount; i++) {
    const rowData = [`Раб. место ${i+1}`];
    let total = 0;
    
    for (const sectionId in checklist.sections) {
      let sectionTotal = 0;
      for (const criteria of checklist.sections[sectionId].criteria) {
        sectionTotal += auditData.results?.[sectionId]?.[criteria.id]?.[i] || 0;
      }
      rowData.push(sectionTotal);
      total += sectionTotal;
    }
    
    rowData.push(total);
    
    const row = worksheet.addRow(rowData);
    row.eachCell(cell => {
      cell.alignment = { horizontal: 'center' };
      cell.border = { left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    });
  }
  
  // Итоговая строка
  const totalRowData = ['Средний балл'];
  for (const sectionId in checklist.sections) {
    let sum = 0;
    for (let i = 0; i < workplaceCount; i++) {
      let sectionTotal = 0;
      for (const criteria of checklist.sections[sectionId].criteria) {
        sectionTotal += auditData.results?.[sectionId]?.[criteria.id]?.[i] || 0;
      }
      sum += sectionTotal;
    }
    totalRowData.push(Math.round(sum / workplaceCount));
  }
  
  const avgTotal = Math.round(totalRowData.slice(1).reduce((a, b) => a + b, 0) / (Object.keys(checklist.sections).length));
  totalRowData.push(avgTotal);
  
  const totalRow = worksheet.addRow(totalRowData);
  totalRow.eachCell(cell => {
    cell.font = { bold: true };
    cell.alignment = { horizontal: 'center' };
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thick' }, right: { style: 'thin' } };
  });
  
  // Настраиваем ширину колонок
  worksheet.getColumn(1).width = 15;
  for (let i = 2; i <= Object.keys(checklist.sections).length + 1; i++) {
    worksheet.getColumn(i).width = 12;
  }
  worksheet.getColumn(Object.keys(checklist.sections).length + 2).width = 10;
}

/**
 * Создать лист с информацией о фото
 */
async function createPhotosSheet(workbook, auditData) {
  const worksheet = workbook.addWorksheet('Фотофиксация');
  
  // Заголовок
  worksheet.getCell('A1').value = 'ФОТОФИКСАЦИЯ ПРОБЛЕМНЫХ ЗОН';
  worksheet.getCell('A1').font = { name: 'Arial', size: 16, bold: true };
  
  worksheet.getCell('A2').value = 'Примечание: Фото сохранены в папке data/photos/';
  worksheet.getCell('A2').font = { name: 'Arial', size: 10, italic: true };
  
  // Заголовки таблицы
  const headers = ['№', 'Рабочее место', 'Секция', 'Критерий', 'Описание', 'Файл'];
  const headerRow = worksheet.addRow(headers);
  headerRow.eachCell(cell => {
    cell.font = { bold: true };
    cell.alignment = { horizontal: 'center' };
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  });
  
  // Данные
  if (auditData.photos && auditData.photos.length > 0) {
    auditData.photos.forEach((photo, index) => {
      const row = worksheet.addRow([
        index + 1,
        photo.workplace || '',
        photo.section || '',
        photo.criteria || '',
        photo.description || '',
        photo.filename || ''
      ]);
      
      row.eachCell(cell => {
        cell.alignment = { vertical: 'top' };
        cell.border = { left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      });
    });
  } else {
    worksheet.addRow(['', '', '', '', 'Фото не добавлены', '']);
  }
  
  // Настраиваем ширину колонок
  worksheet.getColumn(1).width = 5;
  worksheet.getColumn(2).width = 15;
  worksheet.getColumn(3).width = 15;
  worksheet.getColumn(4).width = 15;
  worksheet.getColumn(5).width = 40;
  worksheet.getColumn(6).width = 30;
}

/**
 * Создать шаблон для ручного плана мероприятий
 */
async function createActionPlanTemplate(workbook) {
  const worksheet = workbook.addWorksheet('План мероприятий');
  
  // Заголовок
  worksheet.getCell('A1').value = 'ПЛАН КОРРЕКТИРУЮЩИХ МЕРОПРИЯТИЙ';
  worksheet.getCell('A1').font = { name: 'Arial', size: 16, bold: true };
  worksheet.getCell('A2').value = '(Заполняется вручную после аудита)';
  worksheet.getCell('A2').font = { name: 'Arial', size: 10, italic: true };
  
  // Заголовки таблицы
  const headers = ['№', 'Замечание по культуре производства', 'Мероприятие', 'Цель мероприятия', 'Срок выполнения', 'Ответственный', 'Результат'];
  const headerRow = worksheet.addRow(headers);
  headerRow.eachCell(cell => {
    cell.font = { bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  });
  
  // Пустые строки для заполнения (20 строк)
  for (let i = 0; i < 20; i++) {
    const row = worksheet.addRow([i + 1, '', '', '', '', '', '']);
    row.eachCell(cell => {
      cell.alignment = { vertical: 'top' };
      cell.border = { left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    });
  }
  
  // Подпись
  worksheet.addRow([]);
  worksheet.addRow([]);
  const signatureRow = worksheet.addRow(['Начальник участка / Мастер:', '', '', '', '', '']);
  signatureRow.getCell(1).font = { bold: true };
  signatureRow.getCell(1).colSpan = 6;
  
  // Настраиваем ширину колонок
  worksheet.getColumn(1).width = 5;
  worksheet.getColumn(2).width = 30;
  worksheet.getColumn(3).width = 25;
  worksheet.getColumn(4).width = 25;
  worksheet.getColumn(5).width = 18;
  worksheet.getColumn(6).width = 18;
  worksheet.getColumn(7).width = 15;
}

/**
 * Создать лист с данными для графиков (для ручного создания графиков в Excel)
 */
async function createChartDataSheet(workbook, auditData) {
  const chartSheet = workbook.addWorksheet('Данные для графиков');
  
  // Заголовок
  chartSheet.getCell('A1').value = 'ДАННЫЕ ДЛЯ СОЗДАНИЯ ГРАФИКОВ';
  chartSheet.getCell('A1').font = { name: 'Arial', size: 16, bold: true };
  
  chartSheet.getCell('A2').value = 'Используйте эти данные для создания графиков в Excel вручную';
  chartSheet.getCell('A2').font = { name: 'Arial', size: 10, italic: true };
  
  // Данные для гистограммы
  chartSheet.addRow([]);
  chartSheet.addRow(['Данные для гистограммы:']);
  chartSheet.addRow(['Секция', 'Процент выполнения']);
  
  for (const sectionId in checklist.sections) {
    const score = auditData.sectionScores?.[sectionId]?.percentage || 0;
    chartSheet.addRow([checklist.sections[sectionId].name, score]);
  }
  
  // Данные для радарной диаграммы
  chartSheet.addRow([]);
  chartSheet.addRow(['Данные для радарной диаграммы:']);
  chartSheet.addRow(['Секция', 'Процент']);
  
  for (const sectionId in checklist.sections) {
    const score = auditData.sectionScores?.[sectionId]?.percentage || 0;
    chartSheet.addRow([checklist.sections[sectionId].name, score]);
  }
  
  // Инструкции
  chartSheet.addRow([]);
  chartSheet.addRow(['ИНСТРУКЦИЯ:']);
  chartSheet.addRow(['1. Выделите данные для гистограммы (A4:B9)']);
  chartSheet.addRow(['2. Вставка -> График -> Гистограмма']);
  chartSheet.addRow(['3. Для радарной диаграммы: выделите (A14:B19) -> Вставка -> График -> Лепестковый']);
}

/**
 * Получить уровень по баллам
 */
function getLevelByScore(score) {
  if (score >= 22) return 'Отлично';
  if (score >= 18) return 'Хорошо';
  if (score >= 14) return 'Удовлетворительно';
  if (score >= 10) return 'Неудовлетворительно';
  return 'Критично';
}

/**
 * Получить цвет для уровня
 */
function getLevelColor(level) {
  const colors = {
    'Отлично': { argb: 'FF00FF00' }, // Зеленый
    'Хорошо': { argb: 'FF0000FF' }, // Синий
    'Удовлетворительно': { argb: 'FFFFFF00' }, // Желтый
    'Неудовлетворительно': { argb: 'FFFF9900' }, // Оранжевый
    'Критично': { argb: 'FFFF0000' } // Красный
  };
  return colors[level] || { argb: 'FF000000' };
}

module.exports = {
  generateFullReport
};
