const fs = require('fs');
const path = require('path');
const checklist = require('../config/checklist.json');

// Пути к директориям
const DATA_DIR = path.join(__dirname, '..', 'data');
const AUDITS_DIR = path.join(DATA_DIR, 'audits');
const PHOTOS_DIR = path.join(DATA_DIR, 'photos');

// Создаем директории если их нет
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);
if (!fs.existsSync(AUDITS_DIR)) fs.mkdirSync(AUDITS_DIR);
if (!fs.existsSync(PHOTOS_DIR)) fs.mkdirSync(PHOTOS_DIR);

/**
 * Создать новый аудит
 */
function createAudit(auditData) {
  const auditId = `audit_${Date.now()}`;
  const auditPath = path.join(AUDITS_DIR, `${auditId}.json`);
  
  const audit = {
    id: auditId,
    date: new Date().toISOString(),
    workplace: auditData.workplace,
    auditor: auditData.auditor,
    results: auditData.results || {},
    photos: auditData.photos || [],
    totalScore: auditData.totalScore || 0,
    sectionScores: auditData.sectionScores || {},
    notes: auditData.notes || ''
  };
  
  fs.writeFileSync(auditPath, JSON.stringify(audit, null, 2));
  return audit;
}

/**
 * Сохранить результаты аудита
 */
function saveAuditResults(auditId, results) {
  const auditPath = path.join(AUDITS_DIR, `${auditId}.json`);
  
  if (!fs.existsSync(auditPath)) {
    throw new Error('Аудит не найден');
  }
  
  const audit = JSON.parse(fs.readFileSync(auditPath));
  audit.results = results;
  audit.totalScore = calculateTotalScore(results);
  audit.sectionScores = calculateSectionScores(results);
  audit.updatedAt = new Date().toISOString();
  
  fs.writeFileSync(auditPath, JSON.stringify(audit, null, 2));
  return audit;
}

/**
 * Рассчитать общий балл
 */
function calculateTotalScore(results) {
  let total = 0;
  
  for (const sectionId in results) {
    for (const criteriaId in results[sectionId]) {
      const value = results[sectionId][criteriaId];
      // Если значение - массив (оценки по рабочим местам), суммируем все элементы
      if (Array.isArray(value)) {
        total += value.reduce((sum, v) => sum + (v || 0), 0);
      } else {
        total += value || 0;
      }
    }
  }
  
  return total;
}

/**
 * Рассчитать баллы по секциям
 */
function calculateSectionScores(results) {
  const sectionScores = {};
  
  for (const sectionId in checklist.sections) {
    let score = 0;
    let maxScore = 0;
    
    for (const criteria of checklist.sections[sectionId].criteria) {
      const criteriaId = criteria.id;
      if (results[sectionId] && results[sectionId][criteriaId] !== undefined) {
        const value = results[sectionId][criteriaId];
        // Если значение - массив (оценки по рабочим местам), суммируем все элементы
        if (Array.isArray(value)) {
          score += value.reduce((sum, v) => sum + (v || 0), 0);
        } else {
          score += value || 0;
        }
      }
      maxScore += criteria.weight;
    }
    
    sectionScores[sectionId] = {
      score: score,
      maxScore: maxScore * (results[sectionId]?.[checklist.sections[sectionId].criteria[0]?.id]?.length || 1),
      percentage: Math.round((score / (maxScore * (results[sectionId]?.[checklist.sections[sectionId].criteria[0]?.id]?.length || 1))) * 100)
    };
  }
  
  return sectionScores;
}

/**
 * Получить все аудиты
 */
function getAllAudits() {
  const files = fs.readdirSync(AUDITS_DIR);
  return files.map(file => {
    const audit = JSON.parse(fs.readFileSync(path.join(AUDITS_DIR, file)));
    return audit;
  }).sort((a, b) => new Date(b.date) - new Date(a.date));
}

/**
 * Получить аудит по ID
 */
function getAuditById(auditId) {
  const auditPath = path.join(AUDITS_DIR, `${auditId}.json`);
  if (!fs.existsSync(auditPath)) {
    return null;
  }
  return JSON.parse(fs.readFileSync(auditPath));
}

/**
 * Добавить фото к аудиту
 */
function addPhotoToAudit(auditId, photoData) {
  const auditPath = path.join(AUDITS_DIR, `${auditId}.json`);
  
  if (!fs.existsSync(auditPath)) {
    throw new Error('Аудит не найден');
  }
  
  const audit = JSON.parse(fs.readFileSync(auditPath));
  if (!audit.photos) {
    audit.photos = [];
  }
  
  audit.photos.push(photoData);
  fs.writeFileSync(auditPath, JSON.stringify(audit, null, 2));
  return audit;
}

/**
 * Удалить аудит
 */
function deleteAudit(auditId) {
  const auditPath = path.join(AUDITS_DIR, `${auditId}.json`);
  
  if (fs.existsSync(auditPath)) {
    fs.unlinkSync(auditPath);
    return true;
  }
  return false;
}

/**
 * Получить статистику по всем аудитам
 */
function getAuditStatistics() {
  const audits = getAllAudits();
  
  if (audits.length === 0) {
    return {
      totalAudits: 0,
      avgScore: 0,
      bestScore: 0,
      worstScore: 0,
      sectionStats: {}
    };
  }
  
  const sectionStats = {};
  let totalScore = 0;
  let bestScore = 0;
  let worstScore = 25;
  
  // Инициализация статистики по секциям
  for (const sectionId in checklist.sections) {
    sectionStats[sectionId] = {
      total: 0,
      count: 0,
      avgPercentage: 0
    };
  }
  
  for (const audit of audits) {
    totalScore += audit.totalScore;
    bestScore = Math.max(bestScore, audit.totalScore);
    worstScore = Math.min(worstScore, audit.totalScore);
    
    for (const sectionId in audit.sectionScores) {
      sectionStats[sectionId].total += audit.sectionScores[sectionId].percentage;
      sectionStats[sectionId].count++;
    }
  }
  
  // Рассчет средних значений
  for (const sectionId in sectionStats) {
    if (sectionStats[sectionId].count > 0) {
      sectionStats[sectionId].avgPercentage = Math.round(
        sectionStats[sectionId].total / sectionStats[sectionId].count
      );
    }
  }
  
  return {
    totalAudits: audits.length,
    avgScore: Math.round(totalScore / audits.length),
    bestScore: bestScore,
    worstScore: worstScore,
    sectionStats: sectionStats
  };
}

module.exports = {
  createAudit,
  saveAuditResults,
  calculateTotalScore,
  calculateSectionScores,
  getAllAudits,
  getAuditById,
  addPhotoToAudit,
  deleteAudit,
  getAuditStatistics
};
