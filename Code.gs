// ================================================================
// Vendor Management - Code.gs
// ================================================================
const CONFIG = {
  SHEET_ID: '1yDLx7Yb0dn2MNBBto22oPXKRLDG0SIvicix73d5emho',
  ALLOWED_DOMAIN: 'theporter.in'
};

const SN = {
  HIRING: 'Hiring',
  REQUIREMENT: 'Requirement',
  ACCESS: 'Access',
  DV: 'Data Validation',
  USERBASE: 'User Base',
  LOG: 'Log',
  SETTINGS: 'Settings'
};

const HEADERS = {
  HIRING: [
    'ID', 'Batch ID', 'Req ID', 'Vendor Name', 'Hiring Date', 'Full Name', 'First Name', 'Last Name',
    'Contact Number', 'Alt Contact Number', 'Personal Gmail ID', 'Work Experience', 'Languages',
    'Interview Date', 'Interview By', 'Interview Status', 'Rejected Reason', 'Selected Process',
    'Training Date', 'Training By', 'Training Status', 'Dropout Reason',
    'Certification By', 'Certification Status', 'Not Certified Reason', 'Certified Date',
    'Date of Joining', 'Final Status', 'Added By', 'Added On', 'Updated By', 'Updated On'
  ],
  REQUIREMENT: [
    'Req ID', 'Vendor Name', 'Process', 'Required Language', 'Head count', 'Language Breakdown',
    'Requirement Skill', 'Required TAT', 'Required By', 'Mail Status', 'Mail Sent On', 'Created By', 'Created On'
  ],
  ACCESS: ['Email', 'Name', 'Role', 'Company', 'Added By', 'Added On'],
  DV: ['Languages', 'Vendors', 'Interview Status', 'Selected Process', 'Training Status', 'Certification Status'],
  USERBASE: ['Email', 'Name', 'Role', 'Company', 'Session Date', 'Session Time'],
  LOG: ['Timestamp', 'Email', 'Name', 'Action', 'Details', 'Record ID'],
  SETTINGS: ['Email', 'Theme', 'Zoom', 'Font Family', 'Font Size', 'Updated On']
};

const DV_DEFAULTS = {
  languages: [
    'Assamese', 'Bengali', 'Bodo', 'Dogri', 'Gujarati', 'Hindi', 'Kannada', 'Kashmiri',
    'Konkani', 'Maithili', 'Malayalam', 'Manipuri', 'Marathi', 'Nepali', 'Odia', 'Punjabi',
    'Sanskrit', 'Santali', 'Sindhi', 'Tamil', 'Telugu', 'Urdu', 'English'
  ],
  vendors: ['Degitide', 'Essencea'],
  interviewStatus: ['Selected', 'Rejected', 'Call not answer', 'Not Available Today'],
  selectedProcess: ['Research & survey', 'PTL', 'Kam Process'],
  trainingStatus: ['Completed', 'Dropout'],
  certStatus: ['Certified', 'Not Certified']
};

const DEFAULT_SETTINGS = {
  theme: 'light',
  zoom: 100,
  fontFamily: 'Inter',
  fontSize: 14
};

const ROLE_OPTIONS = ['Admin', 'Supervisor', 'Vendor'];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Vendor Management')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function initApp() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return { ok: false, msg: 'Not authenticated.' };
    const domain = email.split('@')[1] || '';
    if (domain.toLowerCase() !== CONFIG.ALLOWED_DOMAIN) {
      return { ok: false, msg: 'Access restricted to @' + CONFIG.ALLOWED_DOMAIN + ' only.' };
    }

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);

    let name = email.split('@')[0];
    try {
      const profile = People.People.get('people/me', { personFields: 'names' });
      if (profile.names && profile.names[0]) name = profile.names[0].displayName || name;
    } catch (e) {}

    const access = _lookupAccess(ss, email);
    if (!access.found) {
      const sh = ss.getSheetByName(SN.ACCESS);
      const rows = sh.getDataRange().getValues().filter(function(row, idx) {
        return idx > 0 && row[0];
      });
      if (!rows.length) {
        sh.appendRow([email, name, 'Admin', 'Porter', 'system', _nowISO()]);
        _logSession(ss, email, name, 'Admin', 'Porter');
        _writeLog(ss, email, name, 'FIRST_LOGIN', 'Auto created Admin access', '');
        return {
          ok: true,
          email: email,
          name: name,
          role: 'Admin',
          company: 'Porter',
          dv: _getRealtimeFormData(ss, { email: email, role: 'Admin', company: 'Porter' }),
          settings: _getUserSettings(ss, email)
        };
      }
      return { ok: false, msg: 'No access. Contact your Admin.' };
    }

    _logSession(ss, email, name, access.role, access.company);
    _writeLog(ss, email, name, 'LOGIN', 'Logged in', '');
    return {
      ok: true,
      email: email,
      name: name,
      role: access.role,
      company: access.company,
      dv: _getRealtimeFormData(ss, { email: email, role: access.role, company: access.company }),
      settings: _getUserSettings(ss, email)
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function getRealtimeFormData(ctx) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    return { ok: true, data: _getRealtimeFormData(ss, ctx || {}) };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function getDVData() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    return { ok: true, data: _getDVData(ss) };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function getRequirementData(ctx) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const reqSheet = ss.getSheetByName(SN.REQUIREMENT);
    const reqHeaders = _getHeaders(reqSheet);
    const reqRows = _filterRowsByCompany(reqHeaders, _getDataRows(reqSheet), ctx, 'Vendor Name').filter(function(row) {
      return _s(row[0]);
    });

    const hiringSheet = ss.getSheetByName(SN.HIRING);
    const hiringHeaders = _getHeaders(hiringSheet);
    const hiringRows = _filterRowsByCompany(hiringHeaders, _getDataRows(hiringSheet), ctx, 'Vendor Name');
    const metricsByReq = _buildRequirementMetrics(hiringHeaders, hiringRows);

    return {
      ok: true,
      data: reqRows.map(function(row) {
        const obj = _rowToObject(reqHeaders, row);
        obj.breakdown = _parseRequirementBreakdown(obj);
        obj.metrics = metricsByReq[obj['Req ID']] || _emptyRequirementMetrics();
        return obj;
      })
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function saveRequirement(rec, ctx) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    if (!_canManageAllVendors(ctx)) {
      return { ok: false, msg: 'Only Porter users can add requirements.' };
    }

    const breakdown = (rec.breakdown || []).map(function(item) {
      return {
        language: _s(item.language),
        headCount: Number(item.headCount) || 0
      };
    }).filter(function(item) {
      return item.language && item.headCount > 0;
    });

    if (!_s(rec['Vendor Name']) || !_s(rec.Process) || !_s(rec['Required TAT']) || !breakdown.length) {
      return { ok: false, msg: 'Vendor, process, required TAT and at least one language row are required.' };
    }

    const sh = ss.getSheetByName(SN.REQUIREMENT);
    const headers = _getHeaders(sh);
    const nextReqId = _nextNumericValue(_getDataRows(sh), 0, 10000);
    const rowObj = {
      'Req ID': String(nextReqId),
      'Vendor Name': _s(rec['Vendor Name']),
      'Process': _s(rec.Process),
      'Required Language': breakdown.map(function(item) { return item.language + ' x ' + item.headCount; }).join(', '),
      'Head count': String(breakdown.reduce(function(sum, item) { return sum + item.headCount; }, 0)),
      'Language Breakdown': JSON.stringify(breakdown),
      'Requirement Skill': _s(rec['Requirement Skill']),
      'Required TAT': _s(rec['Required TAT']),
      'Required By': _s(rec['Required By']) || Session.getActiveUser().getEmail(),
      'Mail Status': '',
      'Mail Sent On': '',
      'Created By': Session.getActiveUser().getEmail(),
      'Created On': _nowISO()
    };

    sh.appendRow(_objectToRow(headers, rowObj));
    _writeLog(ss, Session.getActiveUser().getEmail(), _ctxName(ctx), 'ADD_REQUIREMENT', 'Added requirement ' + rowObj['Req ID'], rowObj['Req ID']);
    rowObj.breakdown = breakdown;
    rowObj.metrics = _emptyRequirementMetrics();
    return { ok: true, msg: 'Requirement added!', data: rowObj };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function createRequirementMailDraft(payload, ctx) {
  try {
    const to = _s(payload.to);
    if (!to) return { ok: false, msg: 'To email is required.' };

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.REQUIREMENT);
    const headers = _getHeaders(sh);
    const rows = _getDataRows(sh);
    const idx = _findRowIndexByValue(rows, headers.indexOf('Req ID'), _s(payload.reqId));
    if (idx < 0) return { ok: false, msg: 'Requirement not found.' };

    const req = _rowToObject(headers, rows[idx]);
    if (!_canSeeVendor(req['Vendor Name'], ctx)) {
      return { ok: false, msg: 'You do not have access to this requirement.' };
    }

    const breakdown = _parseRequirementBreakdown(req);
    const requesterEmail = Session.getActiveUser().getEmail();
    const requesterName = _ctxName(ctx) || requesterEmail.split('@')[0];
    const subject = _s(payload.subject) || ('Req ID: ' + req['Req ID'] + ' Requirement of Agents for ' + req['Process'] + ' - ' + req['Vendor Name']);
    const greetingName = _formatRecipientNames(to);
    const htmlBody = _buildRequirementMailHtml({
      requirement: req,
      breakdown: breakdown,
      greetingName: greetingName,
      requesterName: requesterName,
      requesterEmail: requesterEmail
    });
    const textBody = _buildRequirementMailText({
      requirement: req,
      breakdown: breakdown,
      greetingName: greetingName,
      requesterName: requesterName,
      requesterEmail: requesterEmail
    });

    const draft = GmailApp.createDraft(to, subject, textBody, {
      cc: _s(payload.cc),
      bcc: _s(payload.bcc),
      htmlBody: htmlBody
    });

    const rowNumber = idx + 2;
    _setCellByHeader(sh, headers, rowNumber, 'Mail Status', 'Draft Created');
    _setCellByHeader(sh, headers, rowNumber, 'Mail Sent On', _nowISO());
    _writeLog(ss, requesterEmail, requesterName, 'REQ_MAIL_DRAFT', 'Created requirement draft for ' + req['Req ID'], req['Req ID']);
    return { ok: true, msg: 'Gmail draft created.', draftId: draft.getId() };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function getHiringData(ctx) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.HIRING);
    const headers = _getHeaders(sh);
    const rows = _filterRowsByCompany(headers, _getDataRows(sh), ctx, 'Vendor Name').filter(function(row) {
      return _s(row[0]);
    });
    return {
      ok: true,
      data: rows.map(function(row) {
        return _rowToObject(headers, row);
      })
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function saveHiringRecord(rec, ctx) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.HIRING);
    const headers = _getHeaders(sh);
    const rows = _getDataRows(sh);
    const email = Session.getActiveUser().getEmail();
    const now = _nowISO();
    const normalized = _normalizeHiringRecord(rec);

    if (!_canSeeVendor(normalized['Vendor Name'], ctx)) {
      return { ok: false, msg: 'You do not have access to this vendor.' };
    }
    const errors = _validateHiringRecord(normalized);
    if (errors.length) return { ok: false, msg: errors.join(' ') };

    normalized['Final Status'] = _finalStatus(normalized);
    normalized['Updated By'] = email;
    normalized['Updated On'] = now;

    const existingIndex = _findRowIndexByValue(rows, headers.indexOf('ID'), normalized['ID']);
    if (!normalized['ID']) {
      normalized['ID'] = _nextHiringId(rows, normalized['Vendor Name']);
      normalized['Added By'] = email;
      normalized['Added On'] = now;
      sh.appendRow(_objectToRow(headers, normalized));
      _writeLog(ss, email, _ctxName(ctx), 'ADD_HIRING', 'Added candidate ' + normalized['Full Name'], normalized['ID']);
      return { ok: true, id: normalized['ID'], msg: 'Candidate added!' };
    }

    if (existingIndex < 0) return { ok: false, msg: 'Record not found.' };
    const current = _rowToObject(headers, rows[existingIndex]);
    if (!_canSeeVendor(current['Vendor Name'], ctx)) {
      return { ok: false, msg: 'You do not have permission to edit this record.' };
    }

    normalized['Added By'] = current['Added By'];
    normalized['Added On'] = current['Added On'];
    const merged = Object.assign({}, current, normalized);
    merged['Final Status'] = _finalStatus(merged);
    sh.getRange(existingIndex + 2, 1, 1, headers.length).setValues([_objectToRow(headers, merged)]);
    _writeLog(ss, email, _ctxName(ctx), 'UPDATE_HIRING', 'Updated candidate ' + merged['Full Name'], merged['ID']);
    return { ok: true, id: merged['ID'], msg: 'Candidate updated!' };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function deleteHiringRecord(id, ctx) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.HIRING);
    const headers = _getHeaders(sh);
    const rows = _getDataRows(sh);
    const idx = _findRowIndexByValue(rows, headers.indexOf('ID'), _s(id));
    if (idx < 0) return { ok: false, msg: 'Record not found.' };
    const rowObj = _rowToObject(headers, rows[idx]);
    if (!_canSeeVendor(rowObj['Vendor Name'], ctx)) {
      return { ok: false, msg: 'You do not have permission to delete this record.' };
    }
    sh.deleteRow(idx + 2);
    _writeLog(ss, Session.getActiveUser().getEmail(), _ctxName(ctx), 'DELETE_HIRING', 'Deleted candidate ' + id, id);
    return { ok: true };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function bulkUpdateHiringRecords(payload, ctx) {
  try {
    const ids = (payload.ids || []).map(_s).filter(Boolean);
    const phase = _s(payload.phase);
    if (!ids.length || !phase) return { ok: false, msg: 'Select candidates and a bulk action first.' };

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.HIRING);
    const headers = _getHeaders(sh);
    const rows = _getDataRows(sh);
    const idCol = headers.indexOf('ID');
    const itemMap = {};
    (payload.items || []).forEach(function(item) {
      itemMap[_s(item.id)] = item || {};
    });

    let updated = 0;
    rows.forEach(function(row, rowIndex) {
      const id = _s(row[idCol]);
      if (!ids.includes(id)) return;
      const current = _rowToObject(headers, row);
      if (!_canSeeVendor(current['Vendor Name'], ctx)) return;

      const common = payload.common || {};
      const item = itemMap[id] || {};
      if (phase === 'Interview') {
        current['Interview Date'] = _s(common.interviewDate);
        current['Interview By'] = _s(common.interviewBy);
        current['Interview Status'] = _s(item.status);
        current['Rejected Reason'] = current['Interview Status'] === 'Rejected' ? _s(item.reason) : '';
        current['Selected Process'] = current['Interview Status'] === 'Selected' ? _s(item.selectedProcess) : '';
      } else if (phase === 'Training') {
        current['Training Date'] = _s(common.trainingDate);
        current['Training By'] = _s(common.trainingBy);
        current['Training Status'] = _s(item.status);
        current['Dropout Reason'] = current['Training Status'] === 'Dropout' ? _s(item.reason) : '';
      } else if (phase === 'Certification') {
        current['Certified Date'] = _s(common.certifiedDate);
        current['Certification By'] = _s(common.certificationBy);
        current['Certification Status'] = _s(item.status);
        current['Not Certified Reason'] = current['Certification Status'] === 'Not Certified' ? _s(item.reason) : '';
      } else if (phase === 'Working from') {
        current['Date of Joining'] = _s(common.workingFromDate);
      }

      current['Final Status'] = _finalStatus(current);
      current['Updated By'] = Session.getActiveUser().getEmail();
      current['Updated On'] = _nowISO();
      rows[rowIndex] = _objectToRow(headers, current);
      updated++;
    });

    if (!updated) return { ok: false, msg: 'No accessible records were updated.' };
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    _writeLog(ss, Session.getActiveUser().getEmail(), _ctxName(ctx), 'BULK_UPDATE_' + phase.toUpperCase().replace(/\s+/g, '_'), 'Updated ' + updated + ' records', ids.join(','));
    return { ok: true, msg: 'Updated ' + updated + ' candidates.', count: updated };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function fetchGoogleSheetPreview(sheetLink) {
  try {
    const spreadsheetId = _extractSpreadsheetId(sheetLink);
    if (!spreadsheetId) return { ok: false, msg: 'Invalid Google Sheet link.' };
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const gidMatch = String(sheetLink || '').match(/[?&#]gid=(\d+)/);
    let sheet = spreadsheet.getSheets()[0];
    if (gidMatch) {
      const match = spreadsheet.getSheets().filter(function(item) {
        return String(item.getSheetId()) === String(gidMatch[1]);
      })[0];
      if (match) sheet = match;
    }
    const values = sheet.getDataRange().getDisplayValues();
    return {
      ok: true,
      sourceType: 'googleSheet',
      sourceName: spreadsheet.getName() + ' / ' + sheet.getName(),
      headers: (values[0] || []).map(_s),
      rows: values.slice(1, 51).filter(function(row) {
        return row.some(function(cell) { return _s(cell); });
      })
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function saveImportedCandidates(payload, ctx) {
  try {
    const records = payload.records || [];
    if (!records.length) return { ok: false, msg: 'No records to upload.' };

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.HIRING);
    const headers = _getHeaders(sh);
    const existingRows = _getDataRows(sh);
    const existingIds = new Set(existingRows.map(function(row) { return _s(row[0]); }).filter(Boolean));
    const idState = _buildIdState(existingRows);
    const email = Session.getActiveUser().getEmail();
    const now = _nowISO();
    const errors = [];
    const newRows = [];

    records.forEach(function(raw, idx) {
      const rec = _normalizeHiringRecord(raw);
      if (!_canManageAllVendors(ctx) && !rec['Vendor Name']) rec['Vendor Name'] = _s(ctx.company);
      if (!_canSeeVendor(rec['Vendor Name'], ctx)) {
        errors.push('Row ' + (idx + 2) + ': Vendor access denied.');
        return;
      }
      const rowErrors = _validateHiringRecord(rec, true);
      if (rec['ID'] && existingIds.has(rec['ID'])) rowErrors.push('ID already exists');
      if (rowErrors.length) {
        errors.push('Row ' + (idx + 2) + ': ' + rowErrors.join(', '));
        return;
      }

      if (!rec['ID']) rec['ID'] = _nextHiringIdFromState(rec['Vendor Name'], idState);
      existingIds.add(rec['ID']);
      rec['Final Status'] = _finalStatus(rec);
      rec['Added By'] = email;
      rec['Added On'] = now;
      rec['Updated By'] = email;
      rec['Updated On'] = now;
      newRows.push(_objectToRow(headers, rec));
    });

    if (errors.length) {
      return { ok: false, msg: 'Validation failed for uploaded data.', errors: errors };
    }

    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
    _writeLog(ss, email, _ctxName(ctx), 'IMPORT_HIRING', 'Imported ' + newRows.length + ' candidates', '');
    return { ok: true, msg: 'Uploaded ' + newRows.length + ' candidates.', count: newRows.length };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function searchPeople(query) {
  try {
    const q = _s(query);
    if (q.length < 2) return { ok: true, results: [] };
    const response = People.People.searchDirectoryPeople({
      query: q,
      readMask: 'names,emailAddresses',
      sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'],
      pageSize: 8
    });
    const results = [];
    if (response.people) {
      response.people.forEach(function(person) {
        const email = (person.emailAddresses || [])[0] && (person.emailAddresses || [])[0].value || '';
        const name = (person.names || [])[0] && (person.names || [])[0].displayName || email;
        if (email) results.push({ email: email, name: name });
      });
    }
    return { ok: true, results: results };
  } catch (e) {
    return { ok: false, results: [], msg: e.message };
  }
}

function getReportData(ctx) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.HIRING);
    const headers = _getHeaders(sh);
    const rows = _filterRowsByCompany(headers, _getDataRows(sh), ctx, 'Vendor Name').filter(function(row) {
      return _s(row[0]);
    }).map(function(row) {
      return _rowToObject(headers, row);
    });

    const vendorBuckets = {};
    rows.forEach(function(row) {
      const vendor = row['Vendor Name'] || 'Unknown';
      if (!vendorBuckets[vendor]) vendorBuckets[vendor] = _newVendorMetrics(vendor);
      _accumulateVendorMetrics(vendorBuckets[vendor], row);
    });

    const vendorRows = Object.keys(vendorBuckets).sort().map(function(vendor) {
      return _finalizeVendorMetrics(vendorBuckets[vendor]);
    });

    const overall = vendorRows.reduce(function(acc, row) {
      acc.totalCandidates += row.totalCandidates;
      acc.interviewSelected += row.interviewSelected;
      acc.trainingCompleted += row.trainingCompleted;
      acc.certified += row.certified;
      acc.workingFrom += row.workingFrom;
      return acc;
    }, {
      totalCandidates: 0,
      interviewSelected: 0,
      trainingCompleted: 0,
      certified: 0,
      workingFrom: 0
    });

    const topVendor = vendorRows.slice().sort(function(a, b) {
      return b.workingFromRate - a.workingFromRate;
    })[0] || null;

    return {
      ok: true,
      data: {
        overall: overall,
        topVendor: topVendor,
        vendorRows: vendorRows
      }
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function getUserSettings() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    return { ok: true, data: _getUserSettings(ss, Session.getActiveUser().getEmail()) };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function saveUserSettings(settings, ctx) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.SETTINGS);
    const headers = _getHeaders(sh);
    const rows = _getDataRows(sh);
    const email = Session.getActiveUser().getEmail();
    const rowObj = {
      'Email': email,
      'Theme': _s(settings.theme) || DEFAULT_SETTINGS.theme,
      'Zoom': String(Number(settings.zoom) || DEFAULT_SETTINGS.zoom),
      'Font Family': _s(settings.fontFamily) || DEFAULT_SETTINGS.fontFamily,
      'Font Size': String(Number(settings.fontSize) || DEFAULT_SETTINGS.fontSize),
      'Updated On': _nowISO()
    };
    const idx = _findRowIndexByValue(rows, headers.indexOf('Email'), email);
    if (idx >= 0) {
      sh.getRange(idx + 2, 1, 1, headers.length).setValues([_objectToRow(headers, rowObj)]);
    } else {
      sh.appendRow(_objectToRow(headers, rowObj));
    }
    _writeLog(ss, email, _ctxName(ctx), 'SAVE_SETTINGS', 'Updated user settings', '');
    return { ok: true, data: _getUserSettings(ss, email) };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function addDVOption(colKey, value, ctx) {
  try {
    const colMap = { languages: 1, vendors: 2, interviewStatus: 3, selectedProcess: 4, trainingStatus: 5, certStatus: 6 };
    const col = colMap[colKey];
    const val = _s(value);
    if (!col || !val) return { ok: false, msg: 'Invalid request.' };
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.DV);
    const existing = sh.getRange(2, col, Math.max(sh.getLastRow() - 1, 1), 1).getValues().flat().map(_s).filter(Boolean);
    if (existing.map(function(item) { return item.toLowerCase(); }).includes(val.toLowerCase())) {
      return { ok: false, msg: 'Value already exists.' };
    }
    sh.getRange(existing.length + 2, col).setValue(val);
    _writeLog(ss, Session.getActiveUser().getEmail(), _ctxName(ctx), 'ADD_DV', 'Added "' + val + '" to ' + colKey, '');
    return { ok: true };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function saveDVAll(newData, ctx) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.DV);
    const lastRow = sh.getLastRow();
    if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 6).clearContent();
    const keys = ['languages', 'vendors', 'interviewStatus', 'selectedProcess', 'trainingStatus', 'certStatus'];
    const maxLength = Math.max.apply(null, keys.map(function(key) {
      return (newData[key] || []).length || 0;
    }).concat([0]));
    if (maxLength > 0) {
      const rows = [];
      for (let i = 0; i < maxLength; i++) {
        rows.push(keys.map(function(key) {
          return newData[key] && newData[key][i] ? newData[key][i] : '';
        }));
      }
      sh.getRange(2, 1, rows.length, 6).setValues(rows);
    }
    _writeLog(ss, Session.getActiveUser().getEmail(), _ctxName(ctx), 'SAVE_DV', 'Saved data validation values', '');
    return { ok: true };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function getAccessList() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.ACCESS);
    const headers = _getHeaders(sh);
    return {
      ok: true,
      data: _getDataRows(sh).filter(function(row) { return _s(row[0]); }).map(function(row) {
        const obj = _rowToObject(headers, row);
        return {
          email: obj['Email'],
          name: obj['Name'],
          role: obj['Role'],
          company: obj['Company'],
          addedBy: obj['Added By'],
          addedOn: obj['Added On']
        };
      })
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function saveAccess(rec, ctx) {
  try {
    const email = _s(rec.email).toLowerCase();
    const role = _s(rec.role);
    const company = _s(rec.company);
    if (!_isEmail(email) || !ROLE_OPTIONS.includes(role) || !company) {
      return { ok: false, msg: 'Enter a valid email, role and company.' };
    }
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.ACCESS);
    const headers = _getHeaders(sh);
    const rows = _getDataRows(sh);
    const rowObj = {
      'Email': email,
      'Name': _s(rec.name),
      'Role': role,
      'Company': company,
      'Added By': Session.getActiveUser().getEmail(),
      'Added On': _nowISO()
    };
    const idx = _findRowIndexByValue(rows, headers.indexOf('Email'), email);
    if (idx >= 0) {
      const current = _rowToObject(headers, rows[idx]);
      rowObj['Added On'] = current['Added On'];
      sh.getRange(idx + 2, 1, 1, headers.length).setValues([_objectToRow(headers, rowObj)]);
      _writeLog(ss, Session.getActiveUser().getEmail(), _ctxName(ctx), 'UPDATE_ACCESS', 'Updated access for ' + email, '');
    } else {
      sh.appendRow(_objectToRow(headers, rowObj));
      _writeLog(ss, Session.getActiveUser().getEmail(), _ctxName(ctx), 'ADD_ACCESS', 'Added access for ' + email, '');
    }
    return { ok: true };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function removeAccess(targetEmail, ctx) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.ACCESS);
    const headers = _getHeaders(sh);
    const rows = _getDataRows(sh);
    const idx = _findRowIndexByValue(rows, headers.indexOf('Email'), _s(targetEmail).toLowerCase());
    if (idx < 0) return { ok: false, msg: 'Not found.' };
    sh.deleteRow(idx + 2);
    _writeLog(ss, Session.getActiveUser().getEmail(), _ctxName(ctx), 'REMOVE_ACCESS', 'Removed access for ' + targetEmail, '');
    return { ok: true };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function getUserBaseData() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.USERBASE);
    const headers = _getHeaders(sh);
    const logs = _getDataRows(sh).filter(function(row) { return _s(row[0]); }).map(function(row) {
      const obj = _rowToObject(headers, row);
      return {
        email: obj['Email'],
        name: obj['Name'],
        role: obj['Role'],
        company: obj['Company'],
        date: obj['Session Date'],
        time: obj['Session Time']
      };
    });
    const unique = {};
    logs.forEach(function(log) {
      if (!unique[log.email]) unique[log.email] = log.role;
    });
    return {
      ok: true,
      data: logs.reverse(),
      summary: {
        total: Object.keys(unique).length,
        admin: Object.keys(unique).filter(function(email) { return unique[email] === 'Admin'; }).length,
        supervisor: Object.keys(unique).filter(function(email) { return unique[email] === 'Supervisor'; }).length,
        vendor: Object.keys(unique).filter(function(email) { return unique[email] === 'Vendor'; }).length
      }
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function getLogData() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    const sh = ss.getSheetByName(SN.LOG);
    const headers = _getHeaders(sh);
    return {
      ok: true,
      data: _getDataRows(sh).filter(function(row) { return _s(row[0]); }).map(function(row) {
        const obj = _rowToObject(headers, row);
        return {
          timestamp: obj['Timestamp'],
          email: obj['Email'],
          name: obj['Name'],
          action: obj['Action'],
          details: obj['Details'],
          recordId: obj['Record ID']
        };
      }).reverse()
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function ensureAllHeaders(ss) {
  Object.keys(SN).forEach(function(key) {
    const sh = ss.getSheetByName(SN[key]) || ss.insertSheet(SN[key]);
    const wanted = HEADERS[key];
    if (wanted) _ensureSheetHeaders(sh, wanted);
  });
  _ensureDVDefaults(ss);
}

function _ensureSheetHeaders(sh, wantedHeaders) {
  const existing = _getHeaders(sh).filter(Boolean);
  const merged = existing.slice();
  wantedHeaders.forEach(function(header) {
    if (merged.indexOf(header) === -1) merged.push(header);
  });
  const finalHeaders = merged.length ? merged : wantedHeaders.slice();
  sh.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
  sh.getRange(1, 1, 1, finalHeaders.length)
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sh.setFrozenRows(1);
}

function _ensureDVDefaults(ss) {
  const sh = ss.getSheetByName(SN.DV);
  const values = sh.getDataRange().getValues();
  if (values.length > 1 && values[1][0]) return;
  const maxLength = Math.max.apply(null, Object.keys(DV_DEFAULTS).map(function(key) {
    return DV_DEFAULTS[key].length;
  }));
  const rows = [];
  for (let i = 0; i < maxLength; i++) {
    rows.push([
      DV_DEFAULTS.languages[i] || '',
      DV_DEFAULTS.vendors[i] || '',
      DV_DEFAULTS.interviewStatus[i] || '',
      DV_DEFAULTS.selectedProcess[i] || '',
      DV_DEFAULTS.trainingStatus[i] || '',
      DV_DEFAULTS.certStatus[i] || ''
    ]);
  }
  if (rows.length) sh.getRange(2, 1, rows.length, 6).setValues(rows);
}

function _lookupAccess(ss, email) {
  const sh = ss.getSheetByName(SN.ACCESS);
  const headers = _getHeaders(sh);
  const rows = _getDataRows(sh);
  const idx = _findRowIndexByValue(rows, headers.indexOf('Email'), _s(email).toLowerCase());
  if (idx < 0) return { found: false };
  const obj = _rowToObject(headers, rows[idx]);
  return { found: true, role: obj['Role'], company: obj['Company'] };
}

function _getRealtimeFormData(ss, ctx) {
  const dv = _getDVData(ss);
  dv.reqIds = _getVisibleReqIds(ss, ctx);
  return dv;
}

function _getDVData(ss) {
  const sh = ss.getSheetByName(SN.DV);
  const rows = sh.getDataRange().getValues();
  const result = {
    languages: [],
    vendors: [],
    interviewStatus: [],
    selectedProcess: [],
    trainingStatus: [],
    certStatus: []
  };
  for (let i = 1; i < rows.length; i++) {
    if (_s(rows[i][0])) result.languages.push(_s(rows[i][0]));
    if (_s(rows[i][1])) result.vendors.push(_s(rows[i][1]));
    if (_s(rows[i][2])) result.interviewStatus.push(_s(rows[i][2]));
    if (_s(rows[i][3])) result.selectedProcess.push(_s(rows[i][3]));
    if (_s(rows[i][4])) result.trainingStatus.push(_s(rows[i][4]));
    if (_s(rows[i][5])) result.certStatus.push(_s(rows[i][5]));
  }
  return result;
}

function _getVisibleReqIds(ss, ctx) {
  const sh = ss.getSheetByName(SN.REQUIREMENT);
  const headers = _getHeaders(sh);
  return _filterRowsByCompany(headers, _getDataRows(sh), ctx, 'Vendor Name').map(function(row) {
    return _s(row[headers.indexOf('Req ID')]);
  }).filter(Boolean);
}

function _buildRequirementMetrics(headers, rows) {
  const reqIdIdx = headers.indexOf('Req ID');
  const interviewStatusIdx = headers.indexOf('Interview Status');
  const trainingStatusIdx = headers.indexOf('Training Status');
  const certificationStatusIdx = headers.indexOf('Certification Status');
  const dojIdx = headers.indexOf('Date of Joining');
  const result = {};

  rows.forEach(function(row) {
    const reqId = _s(row[reqIdIdx]);
    if (!reqId) return;
    if (!result[reqId]) result[reqId] = _emptyRequirementMetrics();
    const item = result[reqId];
    const interviewStatus = _s(row[interviewStatusIdx]);
    const trainingStatus = _s(row[trainingStatusIdx]);
    const certificationStatus = _s(row[certificationStatusIdx]);
    const doj = _s(row[dojIdx]);

    if (interviewStatus) {
      item.intT++;
      if (interviewStatus === 'Selected') item.intS++;
      if (interviewStatus === 'Rejected') item.intR++;
    }
    if (trainingStatus) {
      item.trnT++;
      if (trainingStatus === 'Completed') item.trnC++;
      if (trainingStatus === 'Dropout') item.trnD++;
    }
    if (certificationStatus) {
      item.crtT++;
      if (certificationStatus === 'Certified') item.crtC++;
      if (certificationStatus === 'Not Certified') item.crtN++;
    }
    if (doj) item.wf++;
  });

  return result;
}

function _emptyRequirementMetrics() {
  return { intT: 0, intS: 0, intR: 0, trnT: 0, trnC: 0, trnD: 0, crtT: 0, crtC: 0, crtN: 0, wf: 0 };
}

function _parseRequirementBreakdown(req) {
  const raw = _s(req['Language Breakdown']);
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (parsed && parsed.length) {
        return parsed.map(function(item) {
          return {
            language: _s(item.language || item.label || item.name),
            headCount: Number(item.headCount || item.count || item.value) || 0
          };
        }).filter(function(item) {
          return item.language && item.headCount > 0;
        });
      }
    } catch (e) {}
  }
  const lang = _s(req['Required Language']);
  const headCount = Number(req['Head count']) || 0;
  return lang ? [{ language: lang, headCount: headCount || 1 }] : [];
}

function _buildRequirementMailHtml(payload) {
  const req = payload.requirement;
  const rows = payload.breakdown.map(function(item) {
    return '<tr><td style="padding:8px 12px;border:1px solid #cbd5e1;">' + item.language + '</td><td style="padding:8px 12px;border:1px solid #cbd5e1;text-align:center;">' + item.headCount + '</td></tr>';
  }).join('');
  return [
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#0f172a;line-height:1.6;">',
    '<p>Hi ' + payload.greetingName + ',</p>',
    '<p><strong>Req ID:</strong> ' + req['Req ID'] + '<br><strong>Process:</strong> ' + req['Process'] + '<br><strong>Required By:</strong> ' + req['Required By'] + '<br><strong>Required TAT:</strong> ' + req['Required TAT'] + '</p>',
    '<p>' + (_s(req['Requirement Skill']) || 'Please find the requirement details below.') + '</p>',
    '<p>Please find the language-wise headcount requirement for the process below.</p>',
    '<table style="border-collapse:collapse;margin:12px 0 18px 0;min-width:360px;"><thead><tr><th style="padding:8px 12px;border:1px solid #0f172a;background:#0ea5e9;color:#ffffff;">Required Language</th><th style="padding:8px 12px;border:1px solid #0f172a;background:#0ea5e9;color:#ffffff;">Head Count</th></tr></thead><tbody>' + rows + '</tbody></table>',
    '<p>Thanks &amp; Regards<br>' + payload.requesterName + '<br>' + payload.requesterEmail + '</p>',
    '<p style="font-size:12px;color:#475569;">This mail system is generated via Vendor Management web app.<br>If any query or doubt in Requirement then please reach out to ' + payload.requesterEmail + '.</p>',
    '<p style="font-size:12px;color:#475569;">This system was developed by <a href="https://porter.darwinbox.in/ms/db/profile/view/672572" target="_blank">John Messiah</a>.</p>',
    '</div>'
  ].join('');
}

function _buildRequirementMailText(payload) {
  const req = payload.requirement;
  const lines = [
    'Hi ' + payload.greetingName + ',',
    '',
    'Req ID: ' + req['Req ID'],
    'Process: ' + req['Process'],
    'Required By: ' + req['Required By'],
    'Required TAT: ' + req['Required TAT'],
    '',
    _s(req['Requirement Skill']) || 'Please find the requirement details below.',
    '',
    'Language-wise headcount requirement:'
  ];
  payload.breakdown.forEach(function(item) {
    lines.push('- ' + item.language + ': ' + item.headCount);
  });
  lines.push('', 'Thanks & Regards', payload.requesterName, payload.requesterEmail);
  return lines.join('\n');
}

function _formatRecipientNames(toValue) {
  const firstEmail = _s(toValue).split(',')[0].trim();
  const local = firstEmail.split('@')[0] || 'Team';
  return local.replace(/[._-]+/g, ' ').replace(/\b\w/g, function(chr) { return chr.toUpperCase(); });
}

function _normalizeHiringRecord(rec) {
  const result = {};
  HEADERS.HIRING.forEach(function(header) {
    result[header] = _s(rec[header]);
  });
  if (result['Batch ID']) {
    result['Batch ID'] = String(result['Batch ID']).replace(/^B/i, '');
    result['Batch ID'] = result['Batch ID'] ? 'B' + result['Batch ID'] : '';
  }
  result['Languages'] = (result['Languages'] || '').split(',').map(function(item) {
    return _s(item);
  }).filter(Boolean).join(', ');
  return result;
}

function _validateHiringRecord(rec, isImport) {
  const errors = [];
  if (!_s(rec['Req ID'])) errors.push('Req ID is required');
  if (!_s(rec['Vendor Name'])) errors.push('Vendor Name is required');
  if (!_s(rec['Hiring Date'])) errors.push('Hiring Date is required');
  if (!_s(rec['Full Name'])) errors.push('Full Name is required');
  if (!_s(rec['First Name'])) errors.push('First Name is required');
  if (!_s(rec['Last Name'])) errors.push('Last Name is required');
  if (!_s(rec['Contact Number']) || _s(rec['Contact Number']).replace(/\D/g, '').length !== 10) errors.push('Contact Number must be 10 digits');
  if (_s(rec['Alt Contact Number']) && _s(rec['Alt Contact Number']).replace(/\D/g, '').length !== 10) errors.push('Alt Contact Number must be 10 digits');
  if (!_isEmail(rec['Personal Gmail ID'])) errors.push('Personal Gmail ID must be a valid email');
  if (!_s(rec['Work Experience'])) errors.push('Work Experience is required');
  if (!isImport && !_s(rec['Languages'])) errors.push('Languages are required');
  return errors;
}

function _finalStatus(rec) {
  if (_s(rec['Date of Joining'])) return 'Working from';
  if (_s(rec['Certification Status'])) return _s(rec['Certification Status']);
  if (_s(rec['Training Status'])) return _s(rec['Training Status']);
  if (_s(rec['Interview Status'])) return _s(rec['Interview Status']);
  return '';
}

function _buildIdState(rows) {
  const state = {};
  rows.forEach(function(row) {
    const id = _s(row[0]);
    if (!id) return;
    const prefix = id.replace(/[0-9]/g, '');
    const num = parseInt(id.replace(/\D/g, ''), 10);
    if (!isNaN(num)) state[prefix] = Math.max(state[prefix] || 0, num);
  });
  return state;
}

function _nextHiringId(rows, vendorName) {
  return _nextHiringIdFromState(vendorName, _buildIdState(rows));
}

function _nextHiringIdFromState(vendorName, state) {
  const prefix = _vendorPrefix(vendorName);
  state[prefix] = (state[prefix] || 0) + 1;
  return prefix + state[prefix];
}

function _vendorPrefix(vendorName) {
  const cleaned = _s(vendorName).replace(/[^A-Za-z0-9]/g, '');
  return (cleaned || 'V').charAt(0).toUpperCase();
}

function _newVendorMetrics(vendorName) {
  return {
    vendorName: vendorName,
    totalCandidates: 0,
    interviewSelected: 0,
    interviewRejected: 0,
    trainingCompleted: 0,
    trainingDropout: 0,
    certified: 0,
    notCertified: 0,
    workingFrom: 0
  };
}

function _accumulateVendorMetrics(bucket, row) {
  bucket.totalCandidates++;
  if (_s(row['Interview Status']) === 'Selected') bucket.interviewSelected++;
  if (_s(row['Interview Status']) === 'Rejected') bucket.interviewRejected++;
  if (_s(row['Training Status']) === 'Completed') bucket.trainingCompleted++;
  if (_s(row['Training Status']) === 'Dropout') bucket.trainingDropout++;
  if (_s(row['Certification Status']) === 'Certified') bucket.certified++;
  if (_s(row['Certification Status']) === 'Not Certified') bucket.notCertified++;
  if (_s(row['Date of Joining'])) bucket.workingFrom++;
}

function _finalizeVendorMetrics(bucket) {
  const total = bucket.totalCandidates || 1;
  bucket.interviewSelectionRate = Math.round((bucket.interviewSelected / total) * 100);
  bucket.trainingCompletionRate = Math.round((bucket.trainingCompleted / total) * 100);
  bucket.certificationRate = Math.round((bucket.certified / total) * 100);
  bucket.workingFromRate = Math.round((bucket.workingFrom / total) * 100);
  return bucket;
}

function _getUserSettings(ss, email) {
  const sh = ss.getSheetByName(SN.SETTINGS);
  const headers = _getHeaders(sh);
  const rows = _getDataRows(sh);
  const idx = _findRowIndexByValue(rows, headers.indexOf('Email'), _s(email).toLowerCase());
  if (idx < 0) return Object.assign({}, DEFAULT_SETTINGS);
  const obj = _rowToObject(headers, rows[idx]);
  return {
    theme: _s(obj['Theme']) || DEFAULT_SETTINGS.theme,
    zoom: Number(obj['Zoom']) || DEFAULT_SETTINGS.zoom,
    fontFamily: _s(obj['Font Family']) || DEFAULT_SETTINGS.fontFamily,
    fontSize: Number(obj['Font Size']) || DEFAULT_SETTINGS.fontSize
  };
}

function _filterRowsByCompany(headers, rows, ctx, vendorHeader) {
  const vendorIdx = headers.indexOf(vendorHeader || 'Vendor Name');
  if (vendorIdx < 0 || _canManageAllVendors(ctx)) return rows;
  const company = _s(ctx && ctx.company);
  return rows.filter(function(row) {
    return _s(row[vendorIdx]) === company;
  });
}

function _canManageAllVendors(ctx) {
  return !ctx || !_s(ctx.company) || _s(ctx.company) === 'Porter';
}

function _canSeeVendor(vendorName, ctx) {
  return _canManageAllVendors(ctx) || _s(vendorName) === _s(ctx && ctx.company);
}

function _getHeaders(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(_s);
}

function _getDataRows(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), 1);
  if (lastRow <= 1) return [];
  return sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function _rowToObject(headers, row) {
  const obj = {};
  headers.forEach(function(header, index) {
    obj[header] = _s(row[index]);
  });
  return obj;
}

function _objectToRow(headers, obj) {
  return headers.map(function(header) {
    return obj[header] !== undefined ? obj[header] : '';
  });
}

function _setCellByHeader(sh, headers, rowNumber, header, value) {
  const col = headers.indexOf(header);
  if (col >= 0) sh.getRange(rowNumber, col + 1).setValue(value);
}

function _findRowIndexByValue(rows, colIndex, value) {
  if (colIndex < 0) return -1;
  const target = _s(value).toLowerCase();
  for (let i = 0; i < rows.length; i++) {
    if (_s(rows[i][colIndex]).toLowerCase() === target) return i;
  }
  return -1;
}

function _nextNumericValue(rows, colIndex, startFrom) {
  let max = startFrom || 0;
  rows.forEach(function(row) {
    const num = parseInt(_s(row[colIndex]), 10);
    if (!isNaN(num) && num > max) max = num;
  });
  return max + 1;
}

function _extractSpreadsheetId(input) {
  const text = _s(input);
  if (!text) return '';
  const direct = text.match(/^[a-zA-Z0-9-_]{20,}$/);
  if (direct) return direct[0];
  const fromUrl = text.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return fromUrl ? fromUrl[1] : '';
}

function _logSession(ss, email, name, role, company) {
  try {
    const now = new Date();
    const tz = Session.getScriptTimeZone();
    ss.getSheetByName(SN.USERBASE).appendRow([
      email,
      name,
      role,
      company,
      Utilities.formatDate(now, tz, 'yyyy-MM-dd'),
      Utilities.formatDate(now, tz, 'HH:mm:ss')
    ]);
  } catch (e) {}
}

function _writeLog(ss, email, name, action, details, recordId) {
  try {
    ss.getSheetByName(SN.LOG).appendRow([_nowISO(), email, name || '', action, details || '', recordId || '']);
  } catch (e) {}
}

function _ctxName(ctx) {
  return _s(ctx && ctx.name);
}

function _isEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(_s(value));
}

function _nowISO() {
  return new Date().toISOString();
}

function _s(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    try {
      return value.toISOString();
    } catch (e) {
      return '';
    }
  }
  return String(value).trim();
}
