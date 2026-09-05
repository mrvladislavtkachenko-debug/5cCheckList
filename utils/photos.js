const fs = require('fs');
const path = require('path');
const sharp = require('sharp');

const PHOTOS_DIR = path.join(__dirname, '..', 'data', 'photos');

// Создаем директорию если ее нет
if (!fs.existsSync(PHOTOS_DIR)) {
  fs.mkdirSync(PHOTOS_DIR, { recursive: true });
}

/**
 * Сохранить фото
 * @param {string} base64Image - Изображение в формате base64
 * @param {string} workplace - Рабочее место
 * @param {string} section - Секция (1C, 2C, etc.)
 * @param {string} criteria - Критерий
 * @param {string} description - Описание
 * @returns {Promise<{filename: string, path: string}>}
 */
async function savePhoto(base64Image, workplace, section, criteria, description = '') {
  try {
    // Генерируем уникальное имя файла
    const timestamp = Date.now();
    const filename = `photo_${workplace}_${section}_${criteria}_${timestamp}.jpg`;
    const filepath = path.join(PHOTOS_DIR, filename);
    
    // Удаляем префикс base64
    const base64Data = base64Image.replace(/^data:image\/\w+;base64,/, '');
    const buffer = Buffer.from(base64Data, 'base64');
    
    // Сжимаем изображение
    const compressedBuffer = await sharp(buffer)
      .resize({ width: 1200 }) // Максимальная ширина 1200px
      .jpeg({ quality: 80 }) // Качество 80%
      .toBuffer();
    
    // Сохраняем файл
    fs.writeFileSync(filepath, compressedBuffer);
    
    return {
      success: true,
      filename: filename,
      filepath: filepath,
      workplace: workplace,
      section: section,
      criteria: criteria,
      description: description,
      timestamp: timestamp
    };
  } catch (error) {
    console.error('Ошибка при сохранении фото:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Сохранить фото из файла
 * @param {string} filePath - Путь к файлу
 * @param {string} workplace - Рабочее место
 * @param {string} section - Секция
 * @param {string} criteria - Критерий
 * @param {string} description - Описание
 * @returns {Promise<{filename: string, path: string}>}
 */
async function savePhotoFromFile(filePath, workplace, section, criteria, description = '') {
  try {
    const timestamp = Date.now();
    const ext = path.extname(filePath).toLowerCase();
    const filename = `photo_${workplace}_${section}_${criteria}_${timestamp}${ext}`;
    const destPath = path.join(PHOTOS_DIR, filename);
    
    // Копируем файл
    fs.copyFileSync(filePath, destPath);
    
    // Сжимаем если это изображение
    if (['.jpg', '.jpeg', '.png', '.webp'].includes(ext)) {
      const compressedBuffer = await sharp(destPath)
        .resize({ width: 1200 })
        .jpeg({ quality: 80 })
        .toBuffer();
      
      fs.writeFileSync(destPath, compressedBuffer);
    }
    
    return {
      success: true,
      filename: filename,
      filepath: destPath,
      workplace: workplace,
      section: section,
      criteria: criteria,
      description: description,
      timestamp: timestamp
    };
  } catch (error) {
    console.error('Ошибка при сохранении фото из файла:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Получить список всех фото
 * @returns {Array<{filename: string, filepath: string, metadata: object}>}
 */
function getAllPhotos() {
  const files = fs.readdirSync(PHOTOS_DIR);
  return files.map(filename => {
    const filepath = path.join(PHOTOS_DIR, filename);
    const stats = fs.statSync(filepath);
    
    return {
      filename: filename,
      filepath: filepath,
      size: stats.size,
      createdAt: stats.birthtime,
      updatedAt: stats.mtime
    };
  }).sort((a, b) => b.createdAt - a.createdAt);
}

/**
 * Получить фото по рабочему месту
 * @param {string} workplace - Рабочее место
 * @returns {Array<{filename: string, filepath: string, metadata: object}>}
 */
function getPhotosByWorkplace(workplace) {
  const allPhotos = getAllPhotos();
  return allPhotos.filter(photo => 
    photo.filename.includes(`_${workplace}_`)
  );
}

/**
 * Удалить фото
 * @param {string} filename - Имя файла
 * @returns {boolean}
 */
function deletePhoto(filename) {
  const filepath = path.join(PHOTOS_DIR, filename);
  if (fs.existsSync(filepath)) {
    fs.unlinkSync(filepath);
    return true;
  }
  return false;
}

/**
 * Удалить все фото для определенного аудита
 * @param {string} auditId - ID аудита
 */
function deletePhotosByAudit(auditId) {
  const allPhotos = getAllPhotos();
  allPhotos.forEach(photo => {
    if (photo.filename.includes(`_${auditId}_`)) {
      deletePhoto(photo.filename);
    }
  });
}

/**
 * Копировать фото в отчет
 * @param {string} filename - Имя файла
 * @param {string} targetDir - Целевая директория
 * @returns {Promise<string>}
 */
async function copyPhotoToReport(filename, targetDir) {
  const sourcePath = path.join(PHOTOS_DIR, filename);
  const targetPath = path.join(targetDir, filename);
  
  if (!fs.existsSync(sourcePath)) {
    throw new Error('Фото не найдено');
  }
  
  if (!fs.existsSync(targetDir)) {
    fs.mkdirSync(targetDir, { recursive: true });
  }
  
  fs.copyFileSync(sourcePath, targetPath);
  return targetPath;
}

/**
 * Извлечь метаданные из имени файла
 * @param {string} filename - Имя файла
 * @returns {object}
 */
function extractMetadataFromFilename(filename) {
  // Формат: photo_<workplace>_<section>_<criteria>_<timestamp>.jpg
  const parts = filename.replace('.jpg', '').replace('.jpeg', '').replace('.png', '').split('_');
  
  if (parts.length >= 5) {
    return {
      workplace: parts[1],
      section: parts[2],
      criteria: parts[3],
      timestamp: parts[4]
    };
  }
  
  return {};
}

module.exports = {
  savePhoto,
  savePhotoFromFile,
  getAllPhotos,
  getPhotosByWorkplace,
  deletePhoto,
  deletePhotosByAudit,
  copyPhotoToReport,
  extractMetadataFromFilename,
  PHOTOS_DIR
};
