const APP_TOKEN = 'CHANGE_THIS_TO_A_LONG_RANDOM_TOKEN';
const SCHOOL_NAME = '민족사관고등학교';
const SETTINGS_KEY = 'MINJOK_SCORE_SETTINGS_V2';

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || '').trim();
  if (action === 'settings') {
    return handleSettingsGet_(e);
  }
  return HtmlService.createHtmlOutput(`${SCHOOL_NAME} Apps Script web app`);
}

function doPost(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || '').trim();
    const token = String((e && e.parameter && e.parameter.token) || '');
    if (!APP_TOKEN || APP_TOKEN.indexOf('CHANGE_THIS') === 0 || token !== APP_TOKEN) {
      return postMessageResult_({ type: 'minjok-score-mail-result', ok: false, message: '앱 토큰이 올바르지 않습니다.' });
    }

    const rawPayload = String((e && e.parameter && e.parameter.payload) || '');
    const payload = rawPayload ? JSON.parse(rawPayload) : null;

    if (action === 'saveSettings') {
      validateSettings_(payload);
      PropertiesService.getScriptProperties().setProperty(SETTINGS_KEY, JSON.stringify(payload));
      return postMessageResult_({ type: 'minjok-score-settings-result', ok: true, message: '설정을 저장했습니다.' });
    }

    if (action === 'sendMail' || (!action && payload && payload.recipient)) {
      validateMailPayload_(payload);
      const settings = getStoredSettings_();
      sendScoreEmail_(payload, settings);
      return postMessageResult_({ type: 'minjok-score-mail-result', ok: true, message: '성적표 PDF를 발송했습니다.' });
    }

    return postMessageResult_({ type: 'minjok-score-mail-result', ok: false, message: '알 수 없는 요청입니다.' });
  } catch (error) {
    console.error(error);
    return postMessageResult_({ type: 'minjok-score-mail-result', ok: false, message: '요청을 처리하지 못했습니다.' });
  }
}

function handleSettingsGet_(e) {
  const callback = String((e && e.parameter && e.parameter.callback) || '').trim();
  const payload = {
    ok: true,
    settings: getStoredSettings_()
  };
  const body = callback
    ? `${callback}(${JSON.stringify(payload).replace(/</g, '\\u003c')});`
    : JSON.stringify(payload);
  return ContentService.createTextOutput(body)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function getStoredSettings_() {
  const raw = PropertiesService.getScriptProperties().getProperty(SETTINGS_KEY);
  if (!raw) return defaultSettings_();
  try {
    const parsed = JSON.parse(raw);
    return {
      semesterWeights: normalizeSemesterWeights_(parsed.semesterWeights),
      subjectWeights: normalizeSubjectWeights_(parsed.subjectWeights),
      mailToken: String(parsed.mailToken || ''),
      mailEndpoint: String(parsed.mailEndpoint || '')
    };
  } catch (error) {
    return defaultSettings_();
  }
}

function defaultSettings_() {
  return {
    semesterWeights: { s21: 20, s22: 20, s31: 30, s32: 30 },
    subjectWeights: { korean: 3, english: 3, math: 3, social: 2, science: 2, arts: 1, pe: 1 },
    mailEndpoint: '',
    mailToken: ''
  };
}

function validateSettings_(settings) {
  if (!settings || typeof settings !== 'object') throw new Error('Invalid settings');
  if (!settings.semesterWeights || !settings.subjectWeights) throw new Error('Invalid settings');
}

function validateMailPayload_(payload) {
  if (!payload || !payload.recipient || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(payload.recipient)) {
    throw new Error('Invalid recipient');
  }
  if (!payload.student || !Array.isArray(payload.semesters) || payload.semesters.length !== 4) {
    throw new Error('Invalid score payload');
  }
}

function normalizeSemesterWeights_(weights) {
  const defaults = defaultSettings_().semesterWeights;
  const source = weights || {};
  return {
    s21: numberOrDefault_(source.s21, defaults.s21),
    s22: numberOrDefault_(source.s22, defaults.s22),
    s31: numberOrDefault_(source.s31, defaults.s31),
    s32: numberOrDefault_(source.s32, defaults.s32)
  };
}

function normalizeSubjectWeights_(weights) {
  const defaults = defaultSettings_().subjectWeights;
  const source = weights || {};
  return {
    korean: numberOrDefault_(source.korean, defaults.korean),
    english: numberOrDefault_(source.english, defaults.english),
    math: numberOrDefault_(source.math, defaults.math),
    social: numberOrDefault_(source.social, defaults.social),
    science: numberOrDefault_(source.science, defaults.science),
    arts: numberOrDefault_(source.arts, defaults.arts),
    pe: numberOrDefault_(source.pe, defaults.pe)
  };
}

function numberOrDefault_(value, fallback) {
  const num = Number(value);
  return Number.isFinite(num) && num >= 0 ? num : fallback;
}

function sendScoreEmail_(payload, settings) {
  let tempId = '';
  const title = buildDocumentTitle_(payload);
  try {
    const document = DocumentApp.create(title);
    tempId = document.getId();
    writeReport_(document, payload);
    document.saveAndClose();
    Utilities.sleep(500);

    const pdfName = `${title}.pdf`;
    const pdfBlob = DriveApp.getFileById(tempId)
      .getAs(MimeType.PDF)
      .setName(pdfName);
    const subject = `${SCHOOL_NAME} 1단계 성적표 PDF`;
    GmailApp.sendEmail(payload.recipient, subject, buildPlainTextBody_(payload), {
      attachments: [pdfBlob],
      name: `${SCHOOL_NAME} 성적 계산기`
    });
  } finally {
    if (tempId) {
      try {
        DriveApp.getFileById(tempId).setTrashed(true);
      } catch (cleanupError) {
        // Ignore cleanup errors.
      }
    }
  }
}

function postMessageResult_(result) {
  const serialized = JSON.stringify(result).replace(/</g, '\\u003c');
  const html = `<!doctype html><html><body><script>window.parent.postMessage(${serialized}, '*');</script></body></html>`;
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function buildDocumentTitle_(payload) {
  const name = sanitizeFilePart_(payload.student.name || '학생');
  return `${SCHOOL_NAME}_${name}_1단계_성적표`;
}

function sanitizeFilePart_(value) {
  return String(value).replace(/[\\/:*?"<>|]/g, '_').slice(0, 80);
}

function formatScore_(value) {
  return value === null || value === undefined || value === '' ? '-' : `${Number(value).toFixed(2)}점`;
}

function formatGrade_(value) {
  return value === 'none' ? '없음' : (value || '-');
}

function writeReport_(document, payload) {
  const body = document.getBody();
  body.clear();
  body.appendParagraph(`${SCHOOL_NAME} 2027학년도 1단계 성적표`)
    .setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('입력값을 기준으로 계산한 예상 결과입니다. 실제 전형에서는 모집요강과 학교생활기록부를 확인하세요.');
  body.appendParagraph('');

  body.appendParagraph('지원자 정보').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendTable([
    ['학생 이름', String(payload.student.name || '미입력')],
    ['중학교', String(payload.student.school || '미입력')],
    ['생성 시각', String(payload.generatedAt || '')]
  ]);

  body.appendParagraph('1단계 산출 결과').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendTable([
    ['교과성적', formatScore_(payload.calculation.subjectTotal)],
    ['출결 감점', `-${Number(payload.calculation.deduction || 0).toFixed(2)}점`],
    ['환산 미인정 결석', `${Number(payload.attendance.totalAbsences || 0)}일`],
    ['1단계 예상성적', `${formatScore_(payload.calculation.finalScore)} / ${Number(payload.calculation.maxScore || 40)}점`]
  ]);

  body.appendParagraph('학기별 산출 내역').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const semesterRows = [['학기', '입력 과목', '학기 환산', '반영점수', '성취도']];
  payload.semesters.forEach(semester => {
    const grades = (semester.grades || [])
      .map(grade => `${grade.label}: ${formatGrade_(grade.value)}`)
      .join(', ');
    semesterRows.push([
      String(semester.label || semester.short || ''),
      `${Number(semester.entered || 0)}/7`,
      formatScore_(semester.score),
      formatScore_(semester.contribution),
      grades
    ]);
  });
  body.appendTable(semesterRows);

  body.appendParagraph('출결 입력').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendTable([
    ['미인정 결석', `${Number(payload.attendance.absences || 0)}일`],
    ['미인정 지각', `${Number(payload.attendance.lates || 0)}회`],
    ['미인정 조퇴', `${Number(payload.attendance.earlyLeaves || 0)}회`],
    ['미인정 결과', `${Number(payload.attendance.results || 0)}회`],
    ['지각·조퇴·결과 환산', `${Number(payload.attendance.convertedAbsences || 0)}일`]
  ]);
  body.appendParagraph('');
  body.appendParagraph('※ 본 성적표는 입력값을 바탕으로 한 예상 산출 결과입니다.');
}

function buildPlainTextBody_(payload) {
  return [
    `${SCHOOL_NAME} 2027학년도 1단계 성적표를 첨부합니다.`,
    '',
    `학생 이름: ${payload.student.name || '미입력'}`,
    `중학교: ${payload.student.school || '미입력'}`,
    `1단계 예상성적: ${formatScore_(payload.calculation.finalScore)} / ${Number(payload.calculation.maxScore || 40)}점`,
    '',
    '첨부된 PDF에서 학기별 성취도와 출결 산출 내역을 확인하세요.',
    '※ 본 결과는 입력값을 바탕으로 한 예상 산출이며, 실제 지원 전 모집요강과 학교생활기록부를 대조하세요.'
  ].join('\n');
}
