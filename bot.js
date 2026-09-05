const { Telegraf, Markup, Scenes, session } = require('telegraf');
const path = require('path');
const fs = require('fs');
const chalk = require('chalk');

require('dotenv').config();

const checklist = require('./config/checklist.json');
const { 
  createAudit, 
  saveAuditResults, 
  getAllAudits, 
  getAuditById,
  getAuditStatistics,
  deleteAudit
} = require('./utils/audit');
const { generateFullReport } = require('./utils/excel');
const { 
  savePhotoFromFile, 
  getAllPhotos,
  getPhotosByWorkplace,
  deletePhoto,
  PHOTOS_DIR
} = require('./utils/photos');

// Создаем директории если их нет
const dirs = ['data', 'data/audits', 'data/photos', 'reports', 'utils'];
dirs.forEach(dir => {
  const dirPath = path.join(__dirname, dir);
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
});

// Создаем бота
const BOT_TOKEN = process.env.BOT_TOKEN || '';

if (!BOT_TOKEN) {
  console.error(chalk.red('❌ Ошибка: Не установлен BOT_TOKEN в .env файле'));
  console.log(chalk.yellow('Создайте файл .env с содержимым:'));
  console.log(chalk.yellow('BOT_TOKEN=ваш_токен_от_BotFather'));
  process.exit(1);
}

const bot = new Telegraf(BOT_TOKEN);

// Используем сессии для хранения состояния
bot.use(session());

// Сцена для проведения аудита
const auditScene = new Scenes.BaseScene('audit');

// Хранилище состояния аудита
const auditState = {};

// Начальная клавиатура
const mainKeyboard = Markup.keyboard([
  ['📋 Новый аудит 5С'],
  ['📊 Мои аудиты'],
  ['📈 Статистика'],
  ['🖼️ Фотофиксация'],
  ['❓ Помощь']
]).resize();

// Админ клавиатура (для тестов)
const adminKeyboard = Markup.keyboard([
  ['📋 Новый аудит 5С'],
  ['📊 Мои аудиты'],
  ['📈 Статистика'],
  ['🖼️ Фотофиксация'],
  ['❓ Помощь', '🔄 Сбросить']
]).resize();

// Старт
bot.start((ctx) => {
  const welcomeMessage = `
🌟 <b>Добро пожаловать в бота для аудита 5С!</b> 🌟

Этот бот поможет вам провести аудит системы 5С на овощном участке по методике МУ-52-2024.

<b>🎯 Основные возможности:</b>
• Проведение аудита 5 рабочих мест
• Оценка по 25 критериям 5С
• Фотофиксация проблемных зон
• Генерация Excel отчетов
• Хранение истории аудитов
• Просмотр статистики

<b>📋 Чтобы начать:</b>
Нажмите "Новый аудит 5С" или выберите команду из меню.

<b>💡 Поддержка:</b>
/help - Помощь
/about - О боте
`;
  
  return ctx.replyWithHTML(welcomeMessage, mainKeyboard);
});

// Команда помощь
bot.command('help', (ctx) => {
  const helpMessage = `
<b>📖 РУКОВОДСТВО ПОЛЬЗОВАТЕЛЯ</b>

<b>🎯 Основные команды:</b>
/start - Начать работу
/help - Помощь
/about - О боте

<b>📋 Как провести аудит:</b>
1. Нажмите "Новый аудит 5С"
2. Введите название участка
3. Введите ваше ФИО
4. Укажите количество рабочих мест
5. Оцените каждый критерий для каждого рабочего места
6. Добавьте фотофиксацию (опционально)
7. Сгенерируйте отчет

<b>📊 Как просмотреть результаты:</b>
• Нажмите "Мои аудиты"
• Выберите аудит из списка
• Просмотрите детали или удалите

<b>📈 Статистика:</b>
• Общий балл
• Средние показатели
• Лучшие/худшие результаты

<b>🖼️ Фотофиксация:</b>
• Просмотр всех фото
• Удаление фото

<b>⚠️ Важно:</b>
• Фото сохраняются в папке data/photos/
• Excel отчеты сохраняются в папке reports/
• Все данные хранятся локально на сервере
`;
  
  return ctx.replyWithHTML(helpMessage, mainKeyboard);
});

// Команда о боте
bot.command('about', (ctx) => {
  const aboutMessage = `
<b>🤖 О БОТЕ</b>

<b>Название:</b> 5C Audit Bot
<b>Версия:</b> 1.0.0
<b>Методика:</b> МУ-52-2024 "5С в производстве"
<b>Назначение:</b> Овощной участок (резка овощей)

<b>📊 Структура оценки:</b>
• 5 этапов 5С
• 25 критериев
• 5 рабочих мест
• Максимальный балл: 125

<b>🎯 Уровни оценки:</b>
• ≥90% - Отлично 🟢
• ≥72% - Хорошо 🟣
• ≥56% - Удовлетворительно 🟡
• ≥40% - Неудовлетворительно 🟠
• <40% - Критично 🔴

<b>🔧 Технологии:</b>
• Node.js
• Telegraf
• ExcelJS
• Sharp

<b>📝 Автор:</b> Создан для автоматизации аудита 5С
`;
  
  return ctx.replyWithHTML(aboutMessage, mainKeyboard);
});

// Новый аудит
bot.hears('📋 Новый аудит 5С', async (ctx) => {
  // Сбрасываем состояние
  auditState[ctx.chat.id] = {
    step: 'workplace',
    data: {}
  };
  
  await ctx.reply('📝 Введите название участка:', {
    reply_markup: {
      force_reply: true
    }
  });
  
  return ctx.scene.enter('audit');
});

// Сцена аудита
AuditScene.hears(/.*/, async (ctx) => {
  const state = auditState[ctx.chat.id] || {};
  const text = ctx.message.text;
  
  switch (state.step) {
    case 'workplace':
      state.data.workplace = text;
      state.step = 'auditor';
      auditState[ctx.chat.id] = state;
      await ctx.reply('👤 Введите ваше ФИО:');
      break;
      
    case 'auditor':
      state.data.auditor = text;
      state.step = 'workplaceCount';
      auditState[ctx.chat.id] = state;
      await ctx.reply('🏭 Введите количество рабочих мест (1-10):', {
        reply_markup: {
          force_reply: true
        }
      });
      break;
      
    case 'workplaceCount':
      const count = parseInt(text);
      if (isNaN(count) || count < 1 || count > 10) {
        await ctx.reply('❌ Пожалуйста, введите число от 1 до 10');
        return;
      }
      
      state.data.workplaceCount = count;
      state.step = 'section_1C';
      state.currentSection = '1C';
      state.currentCriteriaIndex = 0;
      state.results = {};
      auditState[ctx.chat.id] = state;
      
      await startSectionAudit(ctx, state);
      break;
      
    case 'section_1C':
    case 'section_2C':
    case 'section_3C':
    case 'section_4C':
    case 'section_5C':
      await processCriteriaAnswer(ctx, state, text);
      break;
      
    case 'photo_description':
      state.photoData.description = text;
      state.step = 'photo_waiting';
      auditState[ctx.chat.id] = state;
      await ctx.reply('📷 Отправьте фото:');
      break;
      
    case 'photo_waiting':
      if (ctx.message.photo) {
        await processPhoto(ctx, state);
      } else {
        await ctx.reply('❌ Пожалуйста, отправьте фото, а не текст');
      }
      break;
      
    default:
      await ctx.reply('❓ Неизвестный шаг. Пожалуйста, начните заново.');
      delete auditState[ctx.chat.id];
      return ctx.scene.leave();
  }
});

// Обработка ответа по критериям
async function processCriteriaAnswer(ctx, state, text) {
  const section = checklist.sections[state.currentSection];
  const criteria = section.criteria[state.currentCriteriaIndex];
  
  // Проверяем, что ответ - это число 0 или 1
  const score = parseInt(text);
  if (isNaN(score) || ![0, 1].includes(score)) {
    await ctx.reply('❌ Пожалуйста, введите 1 (выполнено) или 0 (не выполнено)');
    return;
  }
  
  // Сохраняем оценку для текущего рабочего места
  if (!state.results[state.currentSection]) {
    state.results[state.currentSection] = {};
  }
  
  if (!state.results[state.currentSection][criteria.id]) {
    state.results[state.currentSection][criteria.id] = [];
  }
  
  state.results[state.currentSection][criteria.id].push(score);
  
  // Если это не последнее рабочее место, спрашиваем следующее
  if (state.results[state.currentSection][criteria.id].length < state.data.workplaceCount) {
    state.currentWorkplace = state.results[state.currentSection][criteria.id].length + 1;
    auditState[ctx.chat.id] = state;
    await ctx.reply(`Раб. место ${state.currentWorkplace}: 1 (выполнено) или 0 (не выполнено)?`);
    return;
  }
  
  // Переходим к следующему критерию
  state.currentCriteriaIndex++;
  
  if (state.currentCriteriaIndex < section.criteria.length) {
    auditState[ctx.chat.id] = state;
    await ctx.reply(`\n📌 <b>${section.name}</b>\n\n${section.criteria[state.currentCriteriaIndex].id}: ${section.criteria[state.currentCriteriaIndex].text}\n\nРаб. место 1: 1 или 0?`, { parse_mode: 'HTML' });
    return;
  }
  
  // Переходим к следующей секции
  const sections = Object.keys(checklist.sections);
  const currentIndex = sections.indexOf(state.currentSection);
  
  if (currentIndex < sections.length - 1) {
    state.currentSection = sections[currentIndex + 1];
    state.currentCriteriaIndex = 0;
    auditState[ctx.chat.id] = state;
    await startSectionAudit(ctx, state);
    return;
  }
  
  // Все секции пройдены - завершаем аудит
  await completeAudit(ctx, state);
}

// Начало оценки секции
async function startSectionAudit(ctx, state) {
  const section = checklist.sections[state.currentSection];
  const firstCriteria = section.criteria[0];
  
  await ctx.reply(`\n📌 <b>${section.name}</b>\n${section.description}\n\n${firstCriteria.id}: ${firstCriteria.text}\n\nРаб. место 1: 1 (выполнено) или 0 (не выполнено)?`, { 
    parse_mode: 'HTML',
    reply_markup: {
      force_reply: true
    }
  });
}

// Завершение аудита
async function completeAudit(ctx, state) {
  // Создаем аудит
  const audit = createAudit({
    workplace: state.data.workplace,
    auditor: state.data.auditor,
    workplaceCount: state.data.workplaceCount,
    results: state.results,
    photos: state.photos || []
  });
  
  // Сохраняем результаты
  const savedAudit = saveAuditResults(audit.id, state.results);
  
  // Формируем отчет
  const reportMessage = `
✅ <b>АУДИТ ЗАВЕРШЕН!</b> ✅

<b>📋 Информация:</b>
• Участок: ${savedAudit.workplace}
• Аудитор: ${savedAudit.auditor}
• Рабочих мест: ${savedAudit.workplaceCount}
• Дата: ${new Date(savedAudit.date).toLocaleDateString('ru-RU')}

<b>📊 Результаты:</b>
• Общий балл: ${savedAudit.totalScore}/125
• Уровень: ${getLevelByScore(savedAudit.totalScore)}

<b>📈 По секциям:</b>
`;
  
  let sectionReport = '';
  for (const sectionId in savedAudit.sectionScores) {
    const score = savedAudit.sectionScores[sectionId];
    const level = getLevelByScore(score.score);
    sectionReport += `• ${checklist.sections[sectionId].name}: ${score.score}/${score.maxScore} (${score.percentage}%) - ${level}\n`;
  }
  
  await ctx.replyWithHTML(reportMessage + sectionReport);
  
  // Генерируем Excel отчет
  try {
    const reportPath = await generateFullReport(savedAudit);
    const reportName = path.basename(reportPath);
    
    await ctx.replyWithDocument(
      { source: reportPath },
      { 
        caption: `📄 Excel отчет: ${reportName}\n\nОтчет содержит 7 листов с результатами аудита.`
      }
    );
  } catch (error) {
    await ctx.reply(`❌ Ошибка при генерации отчета: ${error.message}`);
  }
  
  // Очищаем состояние
  delete auditState[ctx.chat.id];
  await ctx.scene.leave();
  
  // Возвращаемся в главное меню
  await ctx.reply('✅ Аудит завершен! Что дальше?', mainKeyboard);
}

// Обработка фото
async function processPhoto(ctx, state) {
  try {
    // Получаем самое большое фото
    const photo = ctx.message.photo.pop();
    const fileId = photo.file_id;
    const filePath = await ctx.telegram.getFile(fileId);
    const fileUrl = `https://api.telegram.org/file/bot${BOT_TOKEN}/${filePath.file_path}`;
    
    // Скачиваем фото
    const response = await fetch(fileUrl);
    const buffer = await response.buffer();
    
    // Сохраняем фото
    const timestamp = Date.now();
    const filename = `photo_${state.photoData.workplace}_${state.photoData.section}_${state.photoData.criteria}_${timestamp}.jpg`;
    const filepath = path.join(PHOTOS_DIR, filename);
    
    fs.writeFileSync(filepath, buffer);
    
    // Добавляем в список фото
    if (!state.photos) {
      state.photos = [];
    }
    
    state.photos.push({
      filename: filename,
      filepath: filepath,
      workplace: state.photoData.workplace,
      section: state.photoData.section,
      criteria: state.photoData.criteria,
      description: state.photoData.description,
      timestamp: timestamp
    });
    
    await ctx.reply(`✅ Фото сохранено: ${filename}`);
    
    // Возвращаемся к оценке критериев
    state.step = state.photoData.returnStep;
    delete state.photoData;
    auditState[ctx.chat.id] = state;
    
    await processCriteriaAnswer(ctx, state, '');
    
  } catch (error) {
    await ctx.reply(`❌ Ошибка при сохранении фото: ${error.message}`);
    state.step = state.photoData.returnStep;
    delete state.photoData;
    auditState[ctx.chat.id] = state;
  }
}

// Просмотр аудитов
bot.hears('📊 Мои аудиты', async (ctx) => {
  const audits = getAllAudits();
  
  if (audits.length === 0) {
    return ctx.reply('📭 У вас пока нет аудитов. Создайте новый аудит!', mainKeyboard);
  }
  
  let message = '<b>📊 ВАШИ АУДИТЫ</b>\n\n';
  
  audits.slice(0, 10).forEach((audit, index) => {
    const date = new Date(audit.date).toLocaleDateString('ru-RU');
    const level = getLevelByScore(audit.totalScore);
    message += `${index + 1}. <b>${audit.workplace}</b> (${date})\n`;
    message += `   Балл: ${audit.totalScore}/125 - ${level}\n\n`;
  });
  
  if (audits.length > 10) {
    message += `... и еще ${audits.length - 10} аудитов\n`;
  }
  
  const keyboard = Markup.keyboard([
    ['🔍 Просмотреть детали'], 
    ['🗑️ Удалить аудит'],
    ['⬅️ Назад']
  ]).resize();
  
  await ctx.replyWithHTML(message, keyboard);
});

// Просмотр деталей аудита
bot.hears('🔍 Просмотреть детали', async (ctx) => {
  const audits = getAllAudits();
  
  if (audits.length === 0) {
    return ctx.reply('📭 Нет аудитов для просмотра', mainKeyboard);
  }
  
  let message = '📋 Выберите аудит для просмотра:\n\n';
  
  audits.slice(0, 10).forEach((audit, index) => {
    const date = new Date(audit.date).toLocaleDateString('ru-RU');
    message += `${index + 1}. ${audit.workplace} (${date}) - ${audit.totalScore}/125\n`;
  });
  
  await ctx.reply(message);
  
  // Сохраняем список аудитов в сессии
  ctx.session.auditsList = audits;
  ctx.session.waitingForAuditNumber = true;
  ctx.session.action = 'view_details';
});

// Удаление аудита
bot.hears('🗑️ Удалить аудит', async (ctx) => {
  const audits = getAllAudits();
  
  if (audits.length === 0) {
    return ctx.reply('📭 Нет аудитов для удаления', mainKeyboard);
  }
  
  let message = '❌ Выберите аудит для удаления:\n\n';
  
  audits.slice(0, 10).forEach((audit, index) => {
    const date = new Date(audit.date).toLocaleDateString('ru-RU');
    message += `${index + 1}. ${audit.workplace} (${date}) - ${audit.totalScore}/125\n`;
  });
  
  await ctx.reply(message);
  
  // Сохраняем список аудитов в сессии
  ctx.session.auditsList = audits;
  ctx.session.waitingForAuditNumber = true;
  ctx.session.action = 'delete_audit';
});

// Обработка выбора номера аудита
bot.on('text', async (ctx) => {
  if (ctx.session.waitingForAuditNumber) {
    const number = parseInt(ctx.message.text);
    const audits = ctx.session.auditsList || [];
    
    if (isNaN(number) || number < 1 || number > audits.length) {
      await ctx.reply('❌ Пожалуйста, введите правильный номер аудита');
      return;
    }
    
    const audit = audits[number - 1];
    
    if (ctx.session.action === 'view_details') {
      await showAuditDetails(ctx, audit);
    } else if (ctx.session.action === 'delete_audit') {
      await deleteAuditInteractive(ctx, audit);
    }
    
    // Сбрасываем состояние
    ctx.session.waitingForAuditNumber = false;
    ctx.session.auditsList = null;
    ctx.session.action = null;
  }
});

// Показать детали аудита
async function showAuditDetails(ctx, audit) {
  let message = `<b>📄 ДЕТАЛИ АУДИТА</b>\n\n`;
  message += `<b>Участок:</b> ${audit.workplace}\n`;
  message += `<b>Аудитор:</b> ${audit.auditor}\n`;
  message += `<b>Дата:</b> ${new Date(audit.date).toLocaleDateString('ru-RU')}\n`;
  message += `<b>Общий балл:</b> ${audit.totalScore}/125\n`;
  message += `<b>Уровень:</b> ${getLevelByScore(audit.totalScore)}\n\n`;
  
  message += `<b>📊 ПО СЕКЦИЯМ:</b>\n`;
  for (const sectionId in audit.sectionScores) {
    const score = audit.sectionScores[sectionId];
    const level = getLevelByScore(score.score);
    message += `• ${checklist.sections[sectionId].name}: ${score.score}/${score.maxScore} (${score.percentage}%) - ${level}\n`;
  }
  
  if (audit.notes) {
    message += `\n<b>📝 Заметки:</b>\n${audit.notes}\n`;
  }
  
  if (audit.photos && audit.photos.length > 0) {
    message += `\n<b>📷 Фотофиксация:</b> ${audit.photos.length} фото\n`;
  }
  
  message += `\n<b>ID аудита:</b> ${audit.id}`;
  
  const keyboard = Markup.keyboard([
    ['📥 Сгенерировать отчет'],
    ['⬅️ Назад']
  ]).resize();
  
  await ctx.replyWithHTML(message, keyboard);
  
  // Сохраняем ID аудита для генерации отчета
  ctx.session.currentAuditId = audit.id;
}

// Удаление аудита
async function deleteAuditInteractive(ctx, audit) {
  await ctx.reply(`⚠️ Вы уверены, что хотите удалить аудит "${audit.workplace}" от ${new Date(audit.date).toLocaleDateString('ru-RU')}?`, {
    reply_markup: {
      inline_keyboard: [
        [{ text: '✅ Да, удалить', callback_data: `delete_${audit.id}` }],
        [{ text: '❌ Отмена', callback_data: 'cancel_delete' }]
      ]
    }
  });
}

// Обработка callback запросов
bot.action(/^delete_(.+)/, async (ctx) => {
  const auditId = ctx.callbackQuery.data.replace('delete_', '');
  
  try {
    const success = deleteAudit(auditId);
    
    if (success) {
      await ctx.editMessageText('✅ Аудит удален');
    } else {
      await ctx.editMessageText('❌ Ошибка при удалении аудита');
    }
  } catch (error) {
    await ctx.editMessageText(`❌ Ошибка: ${error.message}`);
  }
  
  await ctx.reply('Что дальше?', mainKeyboard);
  await ctx.answerCbQuery();
});

bot.action('cancel_delete', async (ctx) => {
  await ctx.editMessageText('❌ Удаление отменено');
  await ctx.reply('Что дальше?', mainKeyboard);
  await ctx.answerCbQuery();
});

// Генерация отчета для выбранного аудита
bot.hears('📥 Сгенерировать отчет', async (ctx) => {
  if (!ctx.session.currentAuditId) {
    return ctx.reply('❌ Пожалуйста, сначала выберите аудит', mainKeyboard);
  }
  
  const audit = getAuditById(ctx.session.currentAuditId);
  
  if (!audit) {
    return ctx.reply('❌ Аудит не найден', mainKeyboard);
  }
  
  await ctx.reply('⏳ Генерирую отчет...');
  
  try {
    const reportPath = await generateFullReport(audit);
    const reportName = path.basename(reportPath);
    
    await ctx.replyWithDocument(
      { source: reportPath },
      { 
        caption: `📄 Excel отчет: ${reportName}\n\nОтчет содержит 7 листов с результатами аудита.`
      }
    );
    
    await ctx.reply('✅ Отчет сгенерирован!', mainKeyboard);
  } catch (error) {
    await ctx.reply(`❌ Ошибка при генерации отчета: ${error.message}`, mainKeyboard);
  }
});

// Статистика
bot.hears('📈 Статистика', async (ctx) => {
  const stats = getAuditStatistics();
  
  let message = '<b>📈 СТАТИСТИКА ПО АУДИТАМ</b>\n\n';
  
  message += `<b>Общее количество:</b> ${stats.totalAudits} аудитов\n\n`;
  
  if (stats.totalAudits > 0) {
    message += `<b>Средний балл:</b> ${stats.avgScore}/125\n`;
    message += `<b>Лучший результат:</b> ${stats.bestScore}/125\n`;
    message += `<b>Худший результат:</b> ${stats.worstScore}/125\n\n`;
    
    message += `<b>Средние показатели по секциям:</b>\n`;
    for (const sectionId in stats.sectionStats) {
      const section = checklist.sections[sectionId];
      const stat = stats.sectionStats[sectionId];
      message += `• ${section.name}: ${stat.avgPercentage}%\n`;
    }
  } else {
    message += '📭 Нет данных для анализа\n';
  }
  
  await ctx.replyWithHTML(message, mainKeyboard);
});

// Фотофиксация
bot.hears('🖼️ Фотофиксация', async (ctx) => {
  const photos = getAllPhotos();
  
  if (photos.length === 0) {
    return ctx.reply('📭 Нет сохраненных фото', mainKeyboard);
  }
  
  let message = '<b>📷 СПИСОК ФОТО</b>\n\n';
  
  photos.slice(0, 10).forEach((photo, index) => {
    const date = new Date(photo.createdAt).toLocaleDateString('ru-RU');
    message += `${index + 1}. ${photo.filename} (${formatFileSize(photo.size)}) - ${date}\n`;
  });
  
  if (photos.length > 10) {
    message += `... и еще ${photos.length - 10} фото\n`;
  }
  
  const keyboard = Markup.keyboard([
    ['🗑️ Удалить фото'],
    ['📁 Открыть папку'],
    ['⬅️ Назад']
  ]).resize();
  
  await ctx.replyWithHTML(message, keyboard);
});

// Удаление фото
bot.hears('🗑️ Удалить фото', async (ctx) => {
  const photos = getAllPhotos();
  
  if (photos.length === 0) {
    return ctx.reply('📭 Нет фото для удаления', mainKeyboard);
  }
  
  let message = '❌ Выберите фото для удаления:\n\n';
  
  photos.slice(0, 10).forEach((photo, index) => {
    message += `${index + 1}. ${photo.filename}\n`;
  });
  
  await ctx.reply(message);
  
  ctx.session.photosList = photos;
  ctx.session.waitingForPhotoNumber = true;
  ctx.session.action = 'delete_photo';
});

// Обработка выбора номера фото
bot.on('text', async (ctx) => {
  if (ctx.session.waitingForPhotoNumber) {
    const number = parseInt(ctx.message.text);
    const photos = ctx.session.photosList || [];
    
    if (isNaN(number) || number < 1 || number > photos.length) {
      await ctx.reply('❌ Пожалуйста, введите правильный номер фото');
      return;
    }
    
    const photo = photos[number - 1];
    
    if (ctx.session.action === 'delete_photo') {
      const success = deletePhoto(photo.filename);
      
      if (success) {
        await ctx.reply(`✅ Фото ${photo.filename} удалено`);
      } else {
        await ctx.reply(`❌ Ошибка при удалении фото`);
      }
    }
    
    ctx.session.waitingForPhotoNumber = false;
    ctx.session.photosList = null;
    ctx.session.action = null;
    
    await ctx.reply('Что дальше?', mainKeyboard);
  }
});

// Открыть папку с фото
bot.hears('📁 Открыть папку', async (ctx) => {
  await ctx.reply(`📁 Папка с фото: ${PHOTOS_DIR}`);
  await ctx.reply('Что дальше?', mainKeyboard);
});

// Помощь
bot.hears('❓ Помощь', (ctx) => {
  return ctx.replyWithHTML(`
<b>📖 КРАТКАЯ ПОМОЩЬ</b>

<b>🎯 Доступные команды:</b>
• /start - Начать
• /help - Помощь
• /about - О боте

<b>📋 Основные функции:</b>
• Новый аудит 5С
• Мои аудиты
• Статистика
• Фотофиксация

<b>❓ Вопросы?</b>
Обратитесь к администратору
`, mainKeyboard);
});

// Назад
bot.hears('⬅️ Назад', (ctx) => {
  return ctx.reply('⬅️ Возврат в главное меню', mainKeyboard);
});

// Обработка ошибок
bot.catch((err, ctx) => {
  console.error('Ошибка:', err);
  return ctx.reply(`❌ Произошла ошибка: ${err.message}`);
});

// Запуск бота
bot.launch().then(() => {
  console.log(chalk.green('✅ Бот запущен!'));
  console.log(chalk.blue(`🤖 Telegram Bot: @${bot.options.username || '5CAuditBot'}`));
  console.log(chalk.gray('='.repeat(50)));
  console.log(chalk.yellow('Для остановки бота нажмите Ctrl+C'));
});

// Обработка сигналов
process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));

// Вспомогательные функции
function getLevelByScore(score) {
  const percentage = (score / 125) * 100;
  if (percentage >= 90) return 'Отлично 🟢';
  if (percentage >= 72) return 'Хорошо 🟣';
  if (percentage >= 56) return 'Удовлетворительно 🟡';
  if (percentage >= 40) return 'Неудовлетворительно 🟠';
  return 'Критично 🔴';
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
}

// Экспорт для тестирования
module.exports = { bot, getLevelByScore, formatFileSize };
