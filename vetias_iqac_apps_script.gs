/**
 * VETIAS IQAC 2026-2027 Master Google Sheet Sync Backend
 * Deployment ID: AKfycbww96RnmXmiaem0rqngQG5rjvySVRdNXkGbVcLHJXV-
 *
 * SETUP INSTRUCTIONS:
 * 1. Open your Master IQAC Google Sheet -> Extensions -> Apps Script.
 * 2. Paste this entire code into Code.gs and click Save.
 * 3. Click Deploy -> Manage deployments (or New deployment) -> Web app:
 *    - Execute as: Me
 *    - Who has access: Anyone within vetias.ac.in (or Anyone)
 * 4. Ensure your active deployment ID matches AKfycbww96RnmXmiaem0rqngQG5rjvySVRdNXkGbVcLHJXV-
 */

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const raw = e && e.postData ? e.postData.contents : '{}';
    const payload = JSON.parse(raw);
    const dept = String(payload.department || '').trim();
    if (!dept) {
      return jsonResponse({ status: 'error', message: 'Department is required.' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = new Date().toISOString();

    // 1. Save full department JSON snapshot in _DEPT_SNAPSHOTS for 2-way restore & Admin Live Consolidation
    let snapSheet = ss.getSheetByName('_DEPT_SNAPSHOTS');
    if (!snapSheet) {
      snapSheet = ss.insertSheet('_DEPT_SNAPSHOTS');
      snapSheet.appendRow(['Department', 'Last_Updated', 'Coordinator', 'HoD', 'PL_Dean', 'JSON_Payload']);
      snapSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    }

    const snapData = snapSheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < snapData.length; i++) {
      if (String(snapData[i][0]).trim().toLowerCase() === dept.toLowerCase()) {
        foundRow = i + 1;
        break;
      }
    }

    const meta = payload.meta || {};
    const rowValues = [
      dept,
      timestamp,
      meta.coordinator || '',
      meta.hod || '',
      meta.dean || '',
      JSON.stringify(payload)
    ];

    if (foundRow > 0) {
      snapSheet.getRange(foundRow, 1, 1, 6).setValues([rowValues]);
    } else {
      snapSheet.appendRow(rowValues);
    }

    // 2. Write flattened Master_Submission_Matrix rows
    if (Array.isArray(payload.summarySheet)) {
      upsertDepartmentRows(
        ss,
        'Master_Submission_Matrix',
        ['Department', 'S_No', 'Domain', 'Particulars', 'Status', 'Faculty_Incharge', 'Record_Count', 'Updated_At'],
        dept,
        payload.summarySheet.map(item => [
          dept,
          item.sno || '',
          item.domain || '',
          item.particulars || '',
          item.status || 'Ongoing',
          item.incharge || '',
          item.count || 0,
          timestamp
        ])
      );
    }

    // 3. Write flattened tabular data into individual consolidated tabs
    const tables = payload.tables || {};
    Object.keys(tables).forEach(tabName => {
      const rows = tables[tabName];
      if (!Array.isArray(rows) || rows.length === 0) return;
      const headers = ['Department', ...Object.keys(rows[0]).filter(k => k !== 'Department')];
      const matrix = rows.map(r => headers.map(h => (h === 'Department' ? dept : (r[h] !== undefined ? r[h] : ''))));
      upsertDepartmentRows(ss, tabName, headers, dept, matrix);
    });

    return jsonResponse({ status: 'ok', department: dept, updatedAt: timestamp });
  } catch (err) {
    return jsonResponse({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = String(params.action || '').trim().toLowerCase();
    const dept = String(params.dept || '').trim();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const snapSheet = ss.getSheetByName('_DEPT_SNAPSHOTS');

    // Admin Dashboard Live Pull: returns all department snapshots
    if (action === 'all') {
      if (!snapSheet) return jsonResponse({ status: 'ok', departments: [] });
      const data = snapSheet.getDataRange().getValues();
      const allDepts = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][5]) {
          try {
            allDepts.push({
              department: data[i][0],
              updatedAt: data[i][1],
              payload: JSON.parse(data[i][5])
            });
          } catch (parseErr) {}
        }
      }
      return jsonResponse({ status: 'ok', count: allDepts.length, departments: allDepts });
    }

    // Single Department Draft Restore
    if (dept) {
      if (!snapSheet) return jsonResponse({ status: 'not_found' });
      const data = snapSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toLowerCase() === dept.toLowerCase()) {
          return jsonResponse({
            status: 'ok',
            updatedAt: data[i][1],
            payload: JSON.parse(data[i][5])
          });
        }
      }
      return jsonResponse({ status: 'not_found' });
    }

    // Health check when opened directly in browser
    return jsonResponse({
      status: 'active',
      service: 'VETIAS IQAC 2026-2027 Live Sync API',
      deploymentId: 'AKfycbww96RnmXmiaem0rqngQG5rjvySVRdNXkGbVcLHJXV-'
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: String(err) });
  }
}

function upsertDepartmentRows(ss, sheetName, headers, dept, newRows) {
  const safeName = sheetName.substring(0, 31);
  let sheet = ss.getSheetByName(safeName);
  if (!sheet) {
    sheet = ss.insertSheet(safeName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
  }

  const existing = sheet.getDataRange().getValues();
  const kept = [existing[0] || headers];
  for (let i = 1; i < existing.length; i++) {
    if (String(existing[i][0]).trim().toLowerCase() !== dept.toLowerCase()) {
      kept.push(existing[i]);
    }
  }
  const finalData = kept.concat(newRows);
  sheet.clearContents();
  if (finalData.length > 0) {
    sheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
