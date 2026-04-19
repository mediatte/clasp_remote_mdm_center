const SECRET_CONFIG = Object.freeze(
  typeof getRuntimeSecretConfig_ === 'function' ? getRuntimeSecretConfig_() : {}
);

const CONFIG = Object.freeze({
  appName: '학생 요청 센터',
  spreadsheetId: SECRET_CONFIG.spreadsheetId || '',
  defaultLocale: 'ko',
  senderEmail: SECRET_CONFIG.senderEmail || '',
  senderName: 'GW 요청센터',
  triggerVersion: '2026-04-19-1',
  sheetNames: Object.freeze({
    request: 'request',
    teachers: 'teachers',
    students: 'students',
    data: 'data',
  }),
  handlers: Object.freeze({
    requestSheetEdit: 'handleRequestSheetEdit',
  }),
  dateFormat: 'yyyy-MM-dd HH:mm:ss',
  statuses: Object.freeze({
    pending: '대기',
    replied: '답변 완료',
    replyError: '답변 오류',
  }),
  requestColumns: Object.freeze({
    requestId: 'Request ID',
    submittedAt: 'Submitted At',
    studentEmail: 'Student Email',
    requestedItem: 'Requested Item',
    reason: 'Reason',
    details: 'Details',
    status: 'Status',
    teacherReplyDraft: 'Teacher Reply Draft',
    notifyStudent: 'Notify Student',
    replySentAt: 'Reply Sent At',
    lastSentReply: 'Last Sent Reply',
    teacherUpdatedAt: 'Teacher Updated At',
    teacherNotifiedAt: 'Teacher Notified At',
    studentDisplay: 'Student Info',
    statusOptions: 'Status Options',
  }),
  teacherColumns: Object.freeze({
    teacherEmail: 'Teacher Email',
    name: 'Name',
    active: 'Active',
  }),
  studentColumns: Object.freeze({
    studentEmail: 'Student Email',
    studentId: 'Student ID',
    name: 'Name',
  }),
});

const REQUEST_HEADERS = Object.values(CONFIG.requestColumns);
const TEACHER_HEADERS = Object.values(CONFIG.teacherColumns);
const STUDENT_HEADERS = Object.values(CONFIG.studentColumns);
const STATUS_OPTION_VALUES = [
  CONFIG.statuses.pending,
  CONFIG.statuses.replied,
  CONFIG.statuses.replyError,
];

function doGet() {
  ensureProjectReady_();

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CONFIG.appName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupRequestApp() {
  const context = ensureProjectReady_();

  return {
    spreadsheetUrl: context.spreadsheet.getUrl(),
    requestSheetUrl: buildSheetUrl_(context.requestSheet),
    teachersSheetUrl: buildSheetUrl_(context.teachersSheet),
    studentsSheetUrl: buildSheetUrl_(context.studentsSheet),
    installedTrigger: true,
  };
}

function getStudentAppModel() {
  const context = ensureProjectReady_();
  const email = getCurrentUserEmail_();
  const studentInfo = email ? getStudentIdentity_(context.studentsSheet, email) : buildEmptyStudentIdentity_();

  return {
    appName: CONFIG.appName,
    defaultLocale: CONFIG.defaultLocale,
    email: email,
    studentId: studentInfo.studentId,
    studentName: studentInfo.studentName,
    studentDisplay: studentInfo.studentDisplay,
    identifiable: Boolean(email),
    identityMessage: email ? '' : buildIdentityMessage_(),
    requests: email ? getStudentRequestsByEmail_(context.requestSheet, email) : [],
  };
}

function submitRequest(payload) {
  const context = ensureProjectReady_();
  const email = requireCurrentUserEmail_();
  const studentInfo = getStudentIdentity_(context.studentsSheet, email);
  const cleanPayload = sanitizeRequestPayload_(payload);
  const requestId = createRequestId_();
  const teacherEmails = getTeacherEmails_(context.teachersSheet);

  context.requestSheet.appendRow([
    requestId,
    new Date(),
    email,
    cleanPayload.requestedItem,
    cleanPayload.reason,
    cleanPayload.details,
    CONFIG.statuses.pending,
    '',
    false,
    '',
    '',
    '',
    '',
    studentInfo.studentDisplay,
  ]);

  SpreadsheetApp.flush();

  const headerMap = getHeaderMap_(context.requestSheet);
  const rowNumber = findRequestRowById_(context.requestSheet, headerMap, requestId);

  if (!rowNumber) {
    throw new Error('요청은 저장되었지만 방금 추가된 행을 다시 찾지 못했습니다.');
  }

  const notifyCell = context.requestSheet.getRange(
    rowNumber,
    headerMap[CONFIG.requestColumns.notifyStudent]
  );
  const teacherNotifiedCell = context.requestSheet.getRange(
    rowNumber,
    headerMap[CONFIG.requestColumns.teacherNotifiedAt]
  );
  notifyCell.insertCheckboxes();
  notifyCell.setValue(false);

  let message = '요청은 저장되었지만 teachers 탭에 활성 교사 이메일이 아직 없습니다.';
  teacherNotifiedCell.setNote('');

  if (teacherEmails.length) {
    try {
      notifyTeachers_(teacherEmails, {
        requestId: requestId,
        studentEmail: email,
        studentDisplay: studentInfo.studentDisplay,
        requestedItem: cleanPayload.requestedItem,
        reason: cleanPayload.reason,
        details: cleanPayload.details,
        sheetUrl: buildSheetRowUrl_(context.requestSheet, rowNumber),
        webAppUrl: getWebAppUrl_(),
      });

      teacherNotifiedCell.setValue(new Date());
      teacherNotifiedCell.setNote('');
      message = '요청이 접수되었고 담당 교사에게 이메일이 발송되었습니다.';
    } catch (error) {
      teacherNotifiedCell.setNote('메일 발송 실패: ' + error.message);
      message = '요청은 저장되었지만 교사 알림 메일 발송에 실패했습니다. ' + error.message;
    }
  } else {
    teacherNotifiedCell.setNote('활성 교사 이메일이 teachers 탭에 없습니다.');
  }

  return {
    message: message,
    requests: getStudentRequestsByEmail_(context.requestSheet, email),
  };
}

function handleRequestSheetEdit(e) {
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  if (normalizeText_(sheet.getName()) !== normalizeText_(CONFIG.sheetNames.request)) {
    return;
  }

  const row = e.range.getRow();
  if (row <= 1) {
    return;
  }

  const headerMap = getHeaderMap_(sheet);
  const replyDraftColumn = headerMap[CONFIG.requestColumns.teacherReplyDraft];
  const notifyStudentColumn = headerMap[CONFIG.requestColumns.notifyStudent];
  const teacherUpdatedAtColumn = headerMap[CONFIG.requestColumns.teacherUpdatedAt];

  if (e.range.getColumn() === replyDraftColumn) {
    sheet.getRange(row, teacherUpdatedAtColumn).setValue(new Date());
    return;
  }

  if (e.range.getColumn() !== notifyStudentColumn || !isCheckedValue_(e.value)) {
    return;
  }

  processTeacherReplyForRow_(sheet, row, headerMap);
}

function processTeacherReplyForRow_(sheet, row, headerMap) {
  const record = getRequestRecordFromRow_(sheet, row, headerMap);
  const notifyCell = sheet.getRange(row, headerMap[CONFIG.requestColumns.notifyStudent]);
  const statusCell = sheet.getRange(row, headerMap[CONFIG.requestColumns.status]);
  const errorTarget = sheet.getRange(row, headerMap[CONFIG.requestColumns.teacherReplyDraft]);

  try {
    if (!isEmail_(record.studentEmail)) {
      throw new Error('이 행에 학생 이메일이 없습니다.');
    }

    if (!record.teacherReplyDraft) {
      throw new Error('Notify Student를 누르기 전에 교사 답변을 입력하세요.');
    }

    if (record.teacherReplyDraft === record.lastSentReply) {
      notifyCell.setValue(false);
      statusCell.setValue(CONFIG.statuses.replied);
      errorTarget.setNote('');
      return;
    }

    sendReplyToStudent_(record, {
      webAppUrl: getWebAppUrl_(),
      sheetUrl: buildSheetRowUrl_(sheet, row),
    });

    sheet.getRange(row, headerMap[CONFIG.requestColumns.replySentAt]).setValue(new Date());
    sheet.getRange(row, headerMap[CONFIG.requestColumns.lastSentReply]).setValue(
      record.teacherReplyDraft
    );
    statusCell.setValue(CONFIG.statuses.replied);
    notifyCell.setValue(false);
    errorTarget.setNote('');
  } catch (error) {
    notifyCell.setValue(false);
    statusCell.setValue(CONFIG.statuses.replyError);
    errorTarget.setNote(error.message);
  }
}

function ensureProjectReady_() {
  assertSecretConfig_();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const requestSheet = getOrCreateSheet_(spreadsheet, CONFIG.sheetNames.request);
    const teachersSheet = getOrCreateSheet_(spreadsheet, CONFIG.sheetNames.teachers);
    const studentLookup = getStudentLookupSheet_(spreadsheet);
    const studentsSheet = studentLookup.sheet;

    ensureSheetHeaders_(requestSheet, REQUEST_HEADERS);
    ensureSheetHeaders_(teachersSheet, TEACHER_HEADERS);
    migrateLegacyRequestColumns_(requestSheet);
    ensureStatusOptions_(requestSheet);
    styleHeaderRow_(requestSheet, REQUEST_HEADERS.length);
    styleHeaderRow_(teachersSheet, TEACHER_HEADERS.length);
    requestSheet.setFrozenRows(1);
    teachersSheet.setFrozenRows(1);

    if (studentLookup.managed) {
      ensureSheetHeaders_(studentsSheet, STUDENT_HEADERS);
      styleHeaderRow_(studentsSheet, STUDENT_HEADERS.length);
      studentsSheet.setFrozenRows(1);
    }

    const requestHeaderMap = getHeaderMap_(requestSheet);
    ensureNotifyCheckboxes_(requestSheet, requestHeaderMap[CONFIG.requestColumns.notifyStudent]);
    ensureStatusValidation_(requestSheet, requestHeaderMap[CONFIG.requestColumns.status]);
    ensureRequestEditTrigger_();

    return {
      spreadsheet: spreadsheet,
      requestSheet: requestSheet,
      teachersSheet: teachersSheet,
      studentsSheet: studentsSheet,
    };
  } finally {
    lock.releaseLock();
  }
}

function ensureRequestEditTrigger_() {
  const properties = PropertiesService.getScriptProperties();
  const versionKey = 'requestSheetEditTriggerVersion';
  const desiredVersion = CONFIG.triggerVersion;
  const matchingTriggers = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return (
      trigger.getHandlerFunction() === CONFIG.handlers.requestSheetEdit &&
      trigger.getTriggerSource &&
      trigger.getTriggerSource() === ScriptApp.TriggerSource.SPREADSHEETS &&
      trigger.getTriggerSourceId &&
      trigger.getTriggerSourceId() === CONFIG.spreadsheetId
    );
  });
  const currentVersion = properties.getProperty(versionKey);
  const needsRefresh = matchingTriggers.length !== 1 || currentVersion !== desiredVersion;

  if (needsRefresh) {
    matchingTriggers.forEach(function(trigger) {
      ScriptApp.deleteTrigger(trigger);
    });

    ScriptApp.newTrigger(CONFIG.handlers.requestSheetEdit)
      .forSpreadsheet(CONFIG.spreadsheetId)
      .onEdit()
      .create();
    properties.setProperty(versionKey, desiredVersion);
  }
}

function getStudentRequestsByEmail_(requestSheet, email) {
  const lastRow = requestSheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const headerMap = getHeaderMap_(requestSheet);
  const values = requestSheet.getRange(2, 1, lastRow - 1, requestSheet.getLastColumn()).getValues();

  return values
    .map(function(row, index) {
      return mapRequestRow_(row, index + 2, headerMap);
    })
    .filter(function(record) {
      return normalizeText_(record.studentEmail) === normalizeText_(email);
    })
    .sort(function(left, right) {
      return new Date(right.submittedAtIso).getTime() - new Date(left.submittedAtIso).getTime();
    })
    .map(function(record) {
      const visibleReply = getVisibleReplyForStudent_(record);

      return {
        requestId: record.requestId,
        submittedAt: record.submittedAtDisplay,
        requestedItem: record.requestedItem,
        reason: record.reason,
        details: record.details,
        status: record.status || CONFIG.statuses.pending,
        studentDisplay: record.studentDisplay,
        reply: visibleReply,
        replySentAt: record.replySentAtDisplay,
      };
    });
}

function getRequestRecordFromRow_(sheet, row, headerMap) {
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  return mapRequestRow_(values, row, headerMap);
}

function mapRequestRow_(row, rowNumber, headerMap) {
  return {
    rowNumber: rowNumber,
    requestId: readRowValue_(row, headerMap[CONFIG.requestColumns.requestId]),
    submittedAt: readRowValue_(row, headerMap[CONFIG.requestColumns.submittedAt]),
    submittedAtIso: toIsoString_(readRowValue_(row, headerMap[CONFIG.requestColumns.submittedAt])),
    submittedAtDisplay: formatDateValue_(readRowValue_(row, headerMap[CONFIG.requestColumns.submittedAt])),
    studentEmail: String(readRowValue_(row, headerMap[CONFIG.requestColumns.studentEmail]) || ''),
    studentDisplay: String(readRowValue_(row, headerMap[CONFIG.requestColumns.studentDisplay]) || ''),
    requestedItem: String(readRowValue_(row, headerMap[CONFIG.requestColumns.requestedItem]) || ''),
    reason: String(readRowValue_(row, headerMap[CONFIG.requestColumns.reason]) || ''),
    details: String(readRowValue_(row, headerMap[CONFIG.requestColumns.details]) || ''),
    status: String(readRowValue_(row, headerMap[CONFIG.requestColumns.status]) || ''),
    teacherReplyDraft: String(
      readRowValue_(row, headerMap[CONFIG.requestColumns.teacherReplyDraft]) || ''
    ),
    notifyStudent: isCheckedValue_(
      readRowValue_(row, headerMap[CONFIG.requestColumns.notifyStudent])
    ),
    replySentAt: readRowValue_(row, headerMap[CONFIG.requestColumns.replySentAt]),
    replySentAtDisplay: formatDateValue_(
      readRowValue_(row, headerMap[CONFIG.requestColumns.replySentAt])
    ),
    teacherUpdatedAt: readRowValue_(row, headerMap[CONFIG.requestColumns.teacherUpdatedAt]),
    lastSentReply: String(readRowValue_(row, headerMap[CONFIG.requestColumns.lastSentReply]) || ''),
  };
}

function getVisibleReplyForStudent_(record) {
  if (record.lastSentReply) {
    return record.lastSentReply;
  }

  if (record.notifyStudent && record.teacherReplyDraft) {
    return record.teacherReplyDraft;
  }

  if (normalizeText_(record.status) === normalizeText_(CONFIG.statuses.replied) && record.teacherReplyDraft) {
    return record.teacherReplyDraft;
  }

  return '';
}

function getTeacherEmails_(teachersSheet) {
  const values = teachersSheet.getDataRange().getValues();
  if (values.length <= 1) {
    return [];
  }

  const normalizedHeaders = values[0].map(normalizeText_);
  const emailColumnIndex = normalizedHeaders.findIndex(function(header) {
    return header.indexOf('email') !== -1;
  });
  const activeColumnIndex = normalizedHeaders.findIndex(function(header) {
    return header === 'active' || header.indexOf('active') !== -1;
  });

  const recipients = new Set();
  values.slice(1).forEach(function(row) {
    const email =
      (emailColumnIndex > -1 ? firstEmailInRow_([row[emailColumnIndex]]) : '') ||
      firstEmailInRow_(row);
    if (!email) {
      return;
    }

    const isInactive =
      activeColumnIndex > -1 && isDisabledFlag_(row[activeColumnIndex]);

    if (!isInactive) {
      recipients.add(email.toLowerCase());
    }
  });

  return Array.from(recipients);
}

function getStudentIdentity_(studentsSheet, email) {
  if (!email) {
    return buildEmptyStudentIdentity_();
  }

  const lastRow = studentsSheet.getLastRow();
  if (lastRow <= 1) {
    return buildDerivedStudentIdentity_(email);
  }

  const headerMap = getFlexibleHeaderMap_(studentsSheet);
  const values = studentsSheet.getRange(2, 1, lastRow - 1, studentsSheet.getLastColumn()).getValues();

  for (let index = 0; index < values.length; index += 1) {
    const row = values[index];
    const candidate = extractStudentIdentityFromRow_(row, headerMap);
    if (normalizeText_(candidate.studentEmail) !== normalizeText_(email)) {
      continue;
    }

    return {
      studentId: candidate.studentId || deriveStudentIdFromEmail_(email),
      studentName: candidate.studentName,
      studentDisplay: buildStudentDisplay_(candidate.studentId || deriveStudentIdFromEmail_(email), candidate.studentName),
    };
  }

  return buildDerivedStudentIdentity_(email);
}

function buildEmptyStudentIdentity_() {
  return {
    studentId: '',
    studentName: '',
    studentDisplay: '',
  };
}

function buildDerivedStudentIdentity_(email) {
  const studentId = deriveStudentIdFromEmail_(email);
  return {
    studentId: studentId,
    studentName: '',
    studentDisplay: buildStudentDisplay_(studentId, ''),
  };
}

function deriveStudentIdFromEmail_(email) {
  const localPart = trimText_(String(email || '').split('@')[0]);
  const match = localPart.match(/\d{4,}/);
  return match ? match[0] : '';
}

function deriveStudentNameFromEmail_(email) {
  const localPart = trimText_(String(email || '').split('@')[0]);
  return localPart
    .replace(/\d+/g, ' ')
    .replace(/[._-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function buildStudentDisplay_(studentId, studentName) {
  const cleanId = trimText_(studentId);
  const cleanName = trimText_(studentName);

  if (cleanName && cleanId && cleanName.indexOf(cleanId) !== -1) {
    return cleanName;
  }

  if (cleanId && cleanName) {
    return cleanId + ' ' + cleanName;
  }

  return cleanName || cleanId || '';
}

function extractStudentIdentityFromRow_(row, headerMap) {
  const mappedEmail = trimText_(readRowValue_(row, headerMap[CONFIG.studentColumns.studentEmail]));
  const mappedId = trimText_(readRowValue_(row, headerMap[CONFIG.studentColumns.studentId]));
  const mappedName = trimText_(readRowValue_(row, headerMap[CONFIG.studentColumns.name]));
  const studentEmail = isEmail_(mappedEmail) ? mappedEmail : firstEmailInRow_(row);
  const studentId =
    isStudentIdLike_(mappedId) ? mappedId : findStudentIdInRow_(row, studentEmail);
  const studentName =
    isStudentNameLike_(mappedName) ? mappedName : findStudentNameInRow_(row, studentEmail, studentId);

  return {
    studentEmail: studentEmail,
    studentId: studentId,
    studentName: studentName,
  };
}

function findStudentIdInRow_(row, studentEmail) {
  for (let index = 0; index < row.length; index += 1) {
    const value = trimText_(row[index]);
    if (!value || value === studentEmail) {
      continue;
    }

    if (isStudentIdLike_(value)) {
      return value;
    }
  }

  return '';
}

function findStudentNameInRow_(row, studentEmail, studentId) {
  for (let index = 0; index < row.length; index += 1) {
    const value = trimText_(row[index]);
    if (!value || value === studentEmail || value === studentId) {
      continue;
    }

    if (isStudentNameLike_(value)) {
      return value;
    }
  }

  return '';
}

function notifyTeachers_(teacherEmails, request) {
  const subject = '[학생 요청] 새 요청: ' + request.requestedItem;
  const plainBody = [
    '학생 요청이 새로 접수되었습니다.',
    '',
    '요청 ID: ' + request.requestId,
    '학생 계정: ' + request.studentEmail,
    '학번/이름: ' + (request.studentDisplay || '(미등록)'),
    '설치 또는 허용 앱/주소: ' + request.requestedItem,
    '사유: ' + request.reason,
    '기타사항: ' + (request.details || '(없음)'),
    '',
    '요청 시트 열기: ' + request.sheetUrl,
    request.webAppUrl ? '학생 페이지 열기: ' + request.webAppUrl : '',
    '',
    '교사는 "Teacher Reply Draft" 열에 답변을 적고 "Notify Student"를 체크하면 학생에게 답변이 발송됩니다.',
  ]
    .filter(Boolean)
    .join('\n');

  const htmlBody = [
    '<p>학생 요청이 새로 접수되었습니다.</p>',
    '<ul>',
    '<li><strong>요청 ID:</strong> ' + escapeHtml_(request.requestId) + '</li>',
    '<li><strong>학생 계정:</strong> ' + escapeHtml_(request.studentEmail) + '</li>',
    '<li><strong>학번/이름:</strong> ' + escapeHtml_(request.studentDisplay || '(미등록)') + '</li>',
    '<li><strong>설치 또는 허용 앱/주소:</strong> ' + escapeHtml_(request.requestedItem) + '</li>',
    '<li><strong>사유:</strong> ' + escapeHtml_(request.reason) + '</li>',
    '<li><strong>기타사항:</strong> ' + escapeHtml_(request.details || '(없음)') + '</li>',
    '</ul>',
    '<p><a href="' + escapeAttribute_(request.sheetUrl) + '">요청 시트 열기</a></p>',
    request.webAppUrl
      ? '<p><a href="' + escapeAttribute_(request.webAppUrl) + '">학생 페이지 열기</a></p>'
      : '',
    '<p><strong>Teacher Reply Draft</strong> 열에 답변을 작성한 뒤 <strong>Notify Student</strong>를 체크하면 학생에게 메일이 발송됩니다.</p>',
  ].join('');

  teacherEmails.forEach(function(email) {
    sendEmailThroughConfiguredSender_({
      to: email,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
      replyTo: CONFIG.senderEmail,
    });
  });
}

function sendReplyToStudent_(record, options) {
  const subject = '[학생 요청 답변] ' + record.requestedItem;
  const plainBody = [
    '학생 요청에 대한 교사 답변이 도착했습니다.',
    '',
    '요청 ID: ' + record.requestId,
    '학번/이름: ' + (record.studentDisplay || '(미등록)'),
    '설치 또는 허용 앱/주소: ' + record.requestedItem,
    '사유: ' + record.reason,
    '',
    '교사 답변:',
    record.teacherReplyDraft,
    '',
    options.webAppUrl ? '학생 페이지 열기: ' + options.webAppUrl : '',
    options.sheetUrl ? '시트 확인: ' + options.sheetUrl : '',
  ]
    .filter(Boolean)
    .join('\n');

  const htmlBody = [
    '<p>학생 요청에 대한 교사 답변이 도착했습니다.</p>',
    '<ul>',
    '<li><strong>요청 ID:</strong> ' + escapeHtml_(record.requestId) + '</li>',
    '<li><strong>학번/이름:</strong> ' + escapeHtml_(record.studentDisplay || '(미등록)') + '</li>',
    '<li><strong>설치 또는 허용 앱/주소:</strong> ' + escapeHtml_(record.requestedItem) + '</li>',
    '<li><strong>사유:</strong> ' + escapeHtml_(record.reason) + '</li>',
    '</ul>',
    '<p><strong>교사 답변</strong></p>',
    '<div style="padding:12px;border-radius:12px;background:#f4f1eb;border:1px solid #ddd6c6;">' +
      escapeHtml_(record.teacherReplyDraft).replace(/\n/g, '<br>') +
      '</div>',
    options.webAppUrl
      ? '<p><a href="' + escapeAttribute_(options.webAppUrl) + '">학생 페이지 열기</a></p>'
      : '',
  ].join('');

  sendEmailThroughConfiguredSender_({
    to: record.studentEmail,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody,
    replyTo: CONFIG.senderEmail,
  });
}

function sanitizeRequestPayload_(payload) {
  const safePayload = payload || {};
  const requestedItem = trimText_(safePayload.requestedItem);
  const reason = trimText_(safePayload.reason);
  const details = trimText_(safePayload.details);

  if (!requestedItem) {
    throw new Error('설치 또는 허용 앱/주소를 입력하세요.');
  }

  if (!reason) {
    throw new Error('사유를 입력하세요.');
  }

  if (requestedItem.length > 200) {
    throw new Error('설치 또는 허용 앱/주소는 200자 이하로 입력하세요.');
  }

  if (reason.length > 1000) {
    throw new Error('사유는 1000자 이하로 입력하세요.');
  }

  if (details.length > 3000) {
    throw new Error('기타사항은 3000자 이하로 입력하세요.');
  }

  return {
    requestedItem: requestedItem,
    reason: reason,
    details: details,
  };
}

function getCurrentUserEmail_() {
  return trimText_(Session.getActiveUser().getEmail());
}

function assertSecretConfig_() {
  const missingKeys = [];

  if (!CONFIG.spreadsheetId) {
    missingKeys.push('APP_SPREADSHEET_ID');
  }

  if (!CONFIG.senderEmail) {
    missingKeys.push('APP_SENDER_EMAIL');
  }

  if (!missingKeys.length) {
    return;
  }

  throw new Error(
    '필수 비공개 설정이 없습니다. .env 기준으로 설정 동기화를 다시 실행하세요: ' +
      missingKeys.join(', ')
  );
}

function sendEmailThroughConfiguredSender_(message) {
  try {
    const quota = MailApp.getRemainingDailyQuota();
    if (quota <= 0) {
      throw new Error('메일 일일 발송 한도를 모두 사용했습니다.');
    }

    MailApp.sendEmail(message.to, message.subject, message.body, {
      htmlBody: message.htmlBody,
      name: CONFIG.senderName || CONFIG.appName,
      replyTo: message.replyTo || CONFIG.senderEmail,
    });
  } catch (error) {
    throw buildMailDeliveryError_(error);
  }
}

function buildMailDeliveryError_(error) {
  const rawMessage = trimText_(error && error.message ? error.message : error);
  const lowerMessage = rawMessage.toLowerCase();

  if (
    lowerMessage.indexOf('authorization') !== -1 ||
    lowerMessage.indexOf('permission') !== -1 ||
    lowerMessage.indexOf('script.send_mail') !== -1 ||
    lowerMessage.indexOf('mail.google.com') !== -1 ||
    lowerMessage.indexOf('gmail.send') !== -1
  ) {
    return new Error(
      '메일 권한 승인이 아직 완료되지 않았습니다. 배포 계정으로 Apps Script 편집기에서 함수 하나를 1회 실행해 메일 권한을 승인하세요. 원본 오류: ' +
        rawMessage
    );
  }

  return error instanceof Error ? error : new Error(rawMessage || '메일 발송 중 오류가 발생했습니다.');
}

function requireCurrentUserEmail_() {
  const email = getCurrentUserEmail_();
  if (!email) {
    throw new Error(buildIdentityMessage_());
  }

  return email;
}

function buildIdentityMessage_() {
  return (
    '현재 로그인한 Google 계정을 확인할 수 없습니다. 도메인 사용자로 웹 앱을 배포해 Session.getActiveUser().getEmail() 값을 읽을 수 있게 해야 합니다.'
  );
}

function getOrCreateSheet_(spreadsheet, desiredName) {
  const existingSheet = spreadsheet.getSheets().find(function(sheet) {
    return normalizeText_(sheet.getName()) === normalizeText_(desiredName);
  });

  return existingSheet || spreadsheet.insertSheet(desiredName);
}

function findSheetByAliases_(spreadsheet, aliases) {
  return spreadsheet.getSheets().find(function(sheet) {
    return aliases.some(function(alias) {
      return normalizeText_(sheet.getName()) === normalizeText_(alias);
    });
  }) || null;
}

function getStudentLookupSheet_(spreadsheet) {
  const studentsSheet = findSheetByAliases_(spreadsheet, [CONFIG.sheetNames.students]);
  const dataSheet = findSheetByAliases_(spreadsheet, [CONFIG.sheetNames.data]);

  if (studentsSheet && studentsSheet.getLastRow() > 1) {
    return {
      sheet: studentsSheet,
      managed: true,
    };
  }

  if (dataSheet && dataSheet.getLastRow() > 1) {
    return {
      sheet: dataSheet,
      managed: false,
    };
  }

  if (studentsSheet) {
    return {
      sheet: studentsSheet,
      managed: true,
    };
  }

  if (dataSheet) {
    return {
      sheet: dataSheet,
      managed: false,
    };
  }

  return {
    sheet: spreadsheet.insertSheet(CONFIG.sheetNames.students),
    managed: true,
  };
}

function ensureSheetHeaders_(sheet, headers) {
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function migrateLegacyRequestColumns_(sheet) {
  const headerMap = getHeaderMap_(sheet);
  const studentDisplayColumn = headerMap[CONFIG.requestColumns.studentDisplay];
  const statusOptionsColumn = headerMap[CONFIG.requestColumns.statusOptions];
  const lastRow = sheet.getLastRow();

  if (studentDisplayColumn && statusOptionsColumn && lastRow > 1) {
    const displayValues = sheet.getRange(2, studentDisplayColumn, lastRow - 1, 1).getValues();
    const adjacentValues = sheet.getRange(2, statusOptionsColumn, lastRow - 1, 1).getValues();
    const mergedValues = displayValues.map(function(displayRow, index) {
      const currentDisplay = trimText_(displayRow[0]);
      const adjacentValue = trimText_(adjacentValues[index][0]);

      if (!adjacentValue || STATUS_OPTION_VALUES.indexOf(adjacentValue) !== -1) {
        return [currentDisplay];
      }

      return [buildStudentDisplay_(currentDisplay, adjacentValue)];
    });

    sheet.getRange(2, studentDisplayColumn, lastRow - 1, 1).setValues(mergedValues);
  }

  if (statusOptionsColumn && sheet.getMaxRows() > 1) {
    sheet.getRange(2, statusOptionsColumn, sheet.getMaxRows() - 1, 1).clearContent();
  }
}

function ensureStatusOptions_(sheet) {
  const headerMap = getHeaderMap_(sheet);
  const statusOptionsColumn = headerMap[CONFIG.requestColumns.statusOptions];

  if (!statusOptionsColumn) {
    return;
  }

  const values = STATUS_OPTION_VALUES.map(function(status) {
    return [status];
  });
  sheet.getRange(2, statusOptionsColumn, values.length, 1).setValues(values);
}

function ensureStatusValidation_(sheet, statusColumn) {
  if (!statusColumn) {
    return;
  }

  const headerMap = getHeaderMap_(sheet);
  const statusOptionsColumn = headerMap[CONFIG.requestColumns.statusOptions];
  if (!statusOptionsColumn) {
    return;
  }

  const optionRange = sheet.getRange(2, statusOptionsColumn, STATUS_OPTION_VALUES.length, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(optionRange, true)
    .setAllowInvalid(false)
    .build();
  const rowCount = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, statusColumn, rowCount, 1).setDataValidation(rule);
}

function styleHeaderRow_(sheet, width) {
  sheet.getRange(1, 1, 1, width)
    .setFontWeight('bold')
    .setBackground('#22313f')
    .setFontColor('#ffffff');
}

function ensureNotifyCheckboxes_(sheet, column) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1 || !column) {
    return;
  }

  sheet.getRange(2, column, lastRow - 1, 1).insertCheckboxes();
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.reduce(function(map, header, index) {
    map[String(header)] = index + 1;
    return map;
  }, {});
}

function getFlexibleHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = headers.reduce(function(accumulator, header, index) {
    accumulator[String(header)] = index + 1;
    return accumulator;
  }, {});
  const normalizedMap = headers.reduce(function(accumulator, header, index) {
    accumulator[normalizeText_(header)] = index + 1;
    return accumulator;
  }, {});

  if (!map[CONFIG.studentColumns.studentEmail]) {
    map[CONFIG.studentColumns.studentEmail] =
      normalizedMap['student email'] ||
      normalizedMap['학생 이메일'] ||
      normalizedMap['email'] ||
      normalizedMap['이메일'] ||
      normalizedMap['메일'] ||
      0;
  }

  if (!map[CONFIG.studentColumns.studentId]) {
    map[CONFIG.studentColumns.studentId] =
      normalizedMap['student id'] ||
      normalizedMap['student number'] ||
      normalizedMap['id'] ||
      normalizedMap['학번'] ||
      normalizedMap['학생 학번'] ||
      0;
  }

  if (!map[CONFIG.studentColumns.name]) {
    map[CONFIG.studentColumns.name] =
      normalizedMap['name'] ||
      normalizedMap['이름'] ||
      normalizedMap['학생 이름'] ||
      normalizedMap['성명'] ||
      0;
  }

  return map;
}

function findRequestRowById_(sheet, headerMap, requestId) {
  const idColumn = headerMap[CONFIG.requestColumns.requestId];
  const lastRow = sheet.getLastRow();

  if (!idColumn || lastRow <= 1) {
    return null;
  }

  const values = sheet.getRange(2, idColumn, lastRow - 1, 1).getValues();

  for (let index = values.length - 1; index >= 0; index -= 1) {
    if (String(values[index][0]) === requestId) {
      return index + 2;
    }
  }

  return null;
}

function buildSheetUrl_(sheet) {
  return sheet.getParent().getUrl() + '#gid=' + sheet.getSheetId();
}

function buildSheetRowUrl_(sheet, row) {
  return buildSheetUrl_(sheet) + '&range=A' + row;
}

function getWebAppUrl_() {
  return ScriptApp.getService().getUrl();
}

function createRequestId_() {
  return 'REQ-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') +
    '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}

function readRowValue_(row, column) {
  return column ? row[column - 1] : '';
}

function formatDateValue_(value) {
  if (!(value instanceof Date)) {
    return value ? String(value) : '';
  }

  return Utilities.formatDate(value, Session.getScriptTimeZone(), CONFIG.dateFormat);
}

function toIsoString_(value) {
  return value instanceof Date ? value.toISOString() : '';
}

function trimText_(value) {
  return String(value || '').trim();
}

function normalizeText_(value) {
  return trimText_(value).toLowerCase();
}

function isCheckedValue_(value) {
  return value === true || normalizeText_(value) === 'true' || normalizeText_(value) === 'yes';
}

function isDisabledFlag_(value) {
  return (
    value === false ||
    ['false', 'no', 'n', '0', 'inactive'].indexOf(normalizeText_(value)) !== -1
  );
}

function isStudentIdLike_(value) {
  const text = trimText_(value);
  if (!text || isEmail_(text)) {
    return false;
  }

  return /^\d{4,}$/.test(text);
}

function isStudentNameLike_(value) {
  const text = trimText_(value);
  if (!text || isEmail_(text) || isStudentIdLike_(text)) {
    return false;
  }

  return true;
}

function firstEmailInRow_(row) {
  for (let index = 0; index < row.length; index += 1) {
    const value = trimText_(row[index]);
    if (isEmail_(value)) {
      return value;
    }
  }

  return '';
}

function isEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimText_(value));
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeAttribute_(value) {
  return escapeHtml_(value);
}
