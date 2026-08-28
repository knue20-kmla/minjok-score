const APP_TOKEN = 'CHANGE_THIS_TO_A_LONG_RANDOM_TOKEN';
const SCHOOL_NAME = '민족사관고등학교';

function doGet() {
  return HtmlService.createHtmlOutput('민족사관고등학교 성적표 메일 발송용 Apps Script가 실행 중입니다.');
}

function doPost(e) {
  let temporaryDocumentId = '';
  try {
    const token = String((e && e.parameter && e.parameter.token) || '');
    if (!APP_TOKEN || APP_TOKEN.indexOf('CHANGE_THIS') === 0 || token !== APP_TOKEN) {
      return response({ ok: false, message: '앱 토큰이 올바르지 않습니다.' });
    }

    const rawPayload = String((e && e.parameter && e.parameter.payload) || '');
    if (!rawPayload) return response({ ok: false, message: '성적표 데이터가 없습니다.' });
    const payload = JSON.parse(rawPayload);
    validatePayload(payload);

    const title = buildDocumentTitle(payload);
    const document = DocumentApp.create(title);
    temporaryDocumentId = document.getId();
    writeReport(document, payload);
    document.saveAndClose();
    Utilities.sleep(500);

    const pdfName = `${title}.pdf`;
    const pdfBlob = DriveApp.getFileById(temporaryDocumentId)
      .getAs(MimeType.PDF)
      .setName(pdfName);
    const subject = `${SCHOOL_NAME} 1단계 성적표 PDF`;
    GmailApp.sendEmail(payload.recipient, subject, buildPlainTextBody(payload), {
      attachments: [pdfBlob],
      name: `${SCHOOL_NAME} 성적 계산기`
    });

    DriveApp.getFileById(temporaryDocumentId).setTrashed(true);
    temporaryDocumentId = '';
    return response({ ok: true, message: '성적표 PDF를 발송했습니다.' });
  } catch (error) {
    if (temporaryDocumentId) {
      try {
        DriveApp.getFileById(temporaryDocumentId).setTrashed(true);
      } catch (cleanupError) {
        // Keep the original error for the client.
      }
    }
    console.error(error);
    return response({ ok: false, message: 'PDF 생성 또는 이메일 발송 중 오류가 발생했습니다.' });
  }
}

function validatePayload(payload) {
  if (!payload || !payload.recipient || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(payload.recipient)) {
    throw new Error('Invalid recipient');
  }
  if (!payload.student || !Array.isArray(payload.semesters) || payload.semesters.length !== 4) {
    throw new Error('Invalid score payload');
  }
}

function response(result) {
  const serialized = JSON.stringify({
    type: 'minjok-score-mail-result',
    ok: Boolean(result.ok),
    message: result.message || ''
  }).replace(/</g, '\\u003c');
  const html = `<!doctype html><html><body><script>window.parent.postMessage(${serialized}, '*');</script></body></html>`;
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function buildDocumentTitle(payload) {
  const name = sanitizeFilePart(payload.student.name || '학생');
  return `${SCHOOL_NAME}_${name}_1단계_성적표`;
}

function sanitizeFilePart(value) {
  return String(value).replace(/[\\/:*?"<>|]/g, '_').slice(0, 80);
}

function formatScore(value) {
  return value === null || value === undefined || value === '' ? '-' : `${Number(value).toFixed(2)}점`;
}

function formatGrade(value) {
  return value === 'none' ? '없음' : (value || '-');
}

function writeReport(document, payload) {
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
    ['교과성적', formatScore(payload.calculation.subjectTotal)],
    ['출결 감점', `-${Number(payload.calculation.deduction || 0).toFixed(2)}점`],
    ['환산 미인정 결석', `${Number(payload.attendance.totalAbsences || 0)}일`],
    ['1단계 예상성적', `${formatScore(payload.calculation.finalScore)} / ${Number(payload.calculation.maxScore || 40)}점`]
  ]);

  body.appendParagraph('학기별 산출 내역').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const semesterRows = [['학기', '입력 과목', '학기 환산', '반영점수', '성취도']];
  payload.semesters.forEach(semester => {
    const grades = (semester.grades || [])
      .map(grade => `${grade.label}: ${formatGrade(grade.value)}`)
      .join(', ');
    semesterRows.push([
      String(semester.label || semester.short || ''),
      `${Number(semester.entered || 0)}/7`,
      formatScore(semester.score),
      formatScore(semester.contribution),
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

function buildPlainTextBody(payload) {
  return [
    `${SCHOOL_NAME} 2027학년도 1단계 성적표를 첨부합니다.`,
    '',
    `학생 이름: ${payload.student.name || '미입력'}`,
    `중학교: ${payload.student.school || '미입력'}`,
    `1단계 예상성적: ${formatScore(payload.calculation.finalScore)} / ${Number(payload.calculation.maxScore || 40)}점`,
    '',
    '첨부된 PDF에서 학기별 성취도와 출결 산출 내역을 확인하세요.',
    '※ 본 결과는 입력값을 바탕으로 한 예상 산출이며, 실제 지원 전 모집요강과 학교생활기록부를 대조하세요.'
  ].join('\n');
}
