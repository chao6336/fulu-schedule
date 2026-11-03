const exportReportNames = {
  tasks: '匯出工作項目',
  personal: '匯出個人班表',
  leaders: '匯出負責人工作表',
  conflicts: '匯出同時段重複工作',
  unassigned: '匯出未安排服務人員'
};

let selectedExportReport = '';

function exportDateKey() {
  const now = new Date();
  return `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
}

function exportTime(value) {
  const raw = String(value ?? '').trim();
  const match = raw.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) return raw;
  return `${match[1].padStart(2, '0')}:${match[2]}:${match[3] || '00'}`;
}

function exportTaskTime(task) {
  return task.allDay ? '整天' : time(task);
}

function exportWeekday(dateValue) {
  const match = String(dateValue ?? '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return '';
  const weekday = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3])).getDay();
  return ['星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'][weekday];
}

function exportText(value) {
  return value === undefined || value === null ? '' : String(value);
}

function exportSort(tasks) {
  return [...tasks].sort((a, b) =>
    taskStartKey(a).localeCompare(taskStartKey(b)) ||
    taskEndKey(a).localeCompare(taskEndKey(b)) ||
    String(a.name || '').localeCompare(String(b.name || ''), 'zh-Hant')
  );
}

function exportTimeMinutes(value) {
  const match = String(value ?? '').match(/^(\d{1,2}):(\d{2})/);
  return match ? Number(match[1]) * 60 + Number(match[2]) : Number.POSITIVE_INFINITY;
}

function exportTimeSeconds(value) {
  const match = String(value ?? '').match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  return match ? Number(match[1]) * 3600 + Number(match[2]) * 60 + Number(match[3] || 0) : Number.POSITIVE_INFINITY;
}

function exportClockFromSeconds(value) {
  const total = Math.max(0, Number(value) || 0);
  const hour = Math.floor(total / 3600) % 24;
  const minute = Math.floor(total % 3600 / 60);
  const second = Math.floor(total % 60);
  return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}:${String(second).padStart(2, '0')}`;
}

function exportDescriptionRowHeight(description) {
  const lines = String(description || '').split(/\r\n|\r|\n/);
  const wrappedLines = lines.reduce((total, line) => {
    const displayWidth = [...line].reduce((width, character) => width + (/\p{Script=Han}/u.test(character) ? 2 : 1), 0);
    return total + Math.max(1, Math.ceil(displayWidth / 52));
  }, 0);
  return Math.min(409, Math.max(22, wrappedLines * 18));
}

function configureExportWorksheet(workbook, name, headers, widths, descriptionColumnIndex) {
  const descriptionColumns = Array.isArray(descriptionColumnIndex)
    ? descriptionColumnIndex
    : descriptionColumnIndex === undefined ? [] : [descriptionColumnIndex];
  const sheet = workbook.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
  sheet.columns = headers.map((header, index) => ({ header, key: `column${index}`, width: widths[index] }));
  sheet.getRow(1).font = { bold: true };
  sheet.getRow(1).alignment = { vertical: 'middle', wrapText: true };
  sheet.getRow(1).height = 24;
  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: headers.length } };
  sheet.columns.forEach((column, index) => {
    column.alignment = descriptionColumns.includes(index)
      ? { vertical: 'top', wrapText: true }
      : { vertical: 'top', wrapText: index === 2 || index === 4 || index === 8 };
  });
  return sheet;
}

function appendExportRows(sheet, rows, descriptionColumnIndex) {
  const descriptionColumns = Array.isArray(descriptionColumnIndex)
    ? descriptionColumnIndex
    : descriptionColumnIndex === undefined ? [] : [descriptionColumnIndex];
  for (const values of rows) {
    const row = sheet.addRow(values.map(value => value === undefined || value === null ? '' : value));
    row.alignment = { vertical: 'top' };
    for (const columnIndex of descriptionColumns) {
      row.getCell(columnIndex + 1).alignment = { vertical: 'top', wrapText: true };
    }
    row.height = descriptionColumns.length
      ? Math.max(...descriptionColumns.map(index => exportDescriptionRowHeight(values[index])))
      : undefined;
  }
}

function taskExportValues(task) {
  const assignedIds = [...new Set(task.assigned || [])];
  const need = Number(task.need || 0);
  const project = (db.projects || []).find(item => String(item.id) === String(task.projectId || ''));
  const owner = (db.projectUsers || []).find(item => String(item.id) === String(task.ownerUserId || ''));
  return [
    exportText(taskStartDate(task)), task.allDay ? '' : exportTime(taskStartTime(task)),
    exportText(taskEndDate(task)), task.allDay ? '' : exportTime(taskEndTime(task)), task.allDay ? '是' : '否',
    exportText(task.name), exportText(task.description),
    exportText(project?.name || '未歸屬專案'), exportText(owner?.name || ''),
    Number.isFinite(Number(task.progress)) ? Number(task.progress) : task.status === '已完成' ? 100 : 0,
    ({ low: '低', normal: '一般', high: '高', urgent: '緊急' })[task.priority] || '一般',
    task.isMilestone ? '是' : '否', task.requiresScheduling === false ? '否' : '是',
    exportText(task.place), need, assignedIds.length, Math.max(0, need - assignedIds.length),
    taskLeaderLabel(task), assignedIds.map(id => personLabel(id)).join('、'), task.status || '規劃中'
  ];
}

function exportWorkItems(workbook) {
  const headers = ['開始日期', '開始時間', '結束日期', '結束時間', '整天', '工作名稱', '工作說明', '專案', '任務負責人', '進度', '優先序', '里程碑', '需要排班', '地點', '需求人數', '已安排人數', '缺額', '主要負責人', '已安排服務人員', '狀態'];
  const widths = [14, 14, 14, 14, 10, 26, 55, 24, 20, 10, 10, 10, 12, 22, 12, 14, 10, 24, 38, 14];
  const sheet = configureExportWorksheet(workbook, '工作項目', headers, widths, 6);
  const rows = exportSort(db.tasks).map(taskExportValues);
  appendExportRows(sheet, rows, 6);
}

function uniqueSheetName(value, used) {
  const base = String(value || '個人班表').replace(/[\\/?*:[\]]/g, '_').slice(0, 31) || '個人班表';
  let name = base;
  let suffix = 2;
  while (used.has(name.toLocaleLowerCase())) {
    const ending = `_${suffix++}`;
    name = `${base.slice(0, 31 - ending.length)}${ending}`;
  }
  used.add(name.toLocaleLowerCase());
  return name;
}

function personalAssignedPeople(task) {
  return (task.assigned || []).map(id => {
    const assignedPerson = db.people.find(personRecord => String(personRecord.id) === String(id));
    return assignedPerson ? assignedPerson.name : personLabel(id);
  }).join('、');
}

function personalRows(personRecord) {
  const assignedTasks = exportSort(db.tasks.filter(task =>
    (task.assigned || []).some(id => String(id) === String(personRecord.id))
  ));
  return assignedTasks.map(task => [
    taskShortDateLabel(task), exportWeekday(taskStartDate(task)), exportTaskTime(task), exportText(task.name),
    exportText(task.description), exportText(task.place), taskLeaderLabel(task),
    personalAssignedPeople(task), task.status || '規劃中'
  ]);
}

function exportPersonalSchedules(workbook, selection) {
  const headers = ['日期', '星期', '時間', '工作名稱', '工作說明', '地點', '主要負責人', '同組服務人員', '狀態'];
  const widths = [14, 13, 24, 26, 55, 22, 24, 24, 14];
  const selectedPeople = selection === 'all'
    ? db.people
    : db.people.filter(personRecord => String(personRecord.id) === selection);
  const people = selectedPeople.length ? selectedPeople : [{ id: '', name: '個人班表' }];
  const usedNames = new Set();
  for (const personRecord of people) {
    const sheet = configureExportWorksheet(workbook, uniqueSheetName(personRecord.name || personRecord.id, usedNames), headers, widths, 4);
    appendExportRows(sheet, personRecord.id ? personalRows(personRecord) : [], 4);
  }
}

function leaderKey(task) {
  return task.leaderId ? `id:${String(task.leaderId)}` : `legacy:${String(task.legacyLeaderName || '')}`;
}

function leaderChoices() {
  const choices = new Map();
  for (const task of db.tasks) {
    const key = leaderKey(task);
    if (key === 'legacy:') continue;
    choices.set(key, taskLeaderLabel(task));
  }
  return [...choices].sort((a, b) => a[1].localeCompare(b[1], 'zh-Hant'));
}

function exportLeaderSheets(workbook, selection) {
  const headers = ['主要負責人', '日期', '星期', '時間', '工作名稱', '工作說明', '地點', '需求人數', '已安排人數', '缺額', '已安排服務人員', '狀態'];
  const widths = [24, 14, 13, 24, 26, 55, 22, 12, 14, 10, 38, 14];
  const sheet = configureExportWorksheet(workbook, '負責人工作表', headers, widths, 5);
  const tasks = exportSort(db.tasks.filter(task => selection === 'all' || leaderKey(task) === selection))
    .sort((a, b) => taskLeaderLabel(a).localeCompare(taskLeaderLabel(b), 'zh-Hant') ||
      taskStartKey(a).localeCompare(taskStartKey(b)) ||
      String(a.name || '').localeCompare(String(b.name || ''), 'zh-Hant'));
  const rows = tasks.map(task => {
    const values = taskExportValues(task);
    return [taskLeaderLabel(task), taskShortDateLabel(task), exportWeekday(taskStartDate(task)), exportTaskTime(task), values[5], values[6], values[13], values[14], values[15], values[16], values[18], values[19]];
  });
  appendExportRows(sheet, rows, 5);
}

function exportRangeValues() {
  return {
    dateFrom: document.querySelector('#exportDateFrom').value,
    dateTo: document.querySelector('#exportDateTo').value,
    timeFrom: document.querySelector('#exportTimeFrom').value,
    timeTo: document.querySelector('#exportTimeTo').value
  };
}

function exportRangeIsValid(range) {
  if (range.dateFrom && range.dateTo && range.dateFrom > range.dateTo) {
    toast('結束日期不可早於起始日期');
    return false;
  }
  const from = range.timeFrom ? exportTimeSeconds(range.timeFrom) : 0;
  const to = range.timeTo ? exportTimeSeconds(range.timeTo) : 86400;
  if ((range.timeFrom || range.timeTo) && from >= to) {
    toast('結束時間必須晚於開始時間');
    return false;
  }
  return true;
}

function taskMatchesExportRange(task, range) {
  if (range.dateFrom && taskEndDate(task) < range.dateFrom) return false;
  if (range.dateTo && taskStartDate(task) > range.dateTo) return false;
  return taskMatchesTime(task, range.timeFrom, range.timeTo);
}

function exportOverlappingAssignments(workbook, range) {
  const headers = ['服務人員', '日期', '重疊時段', '工作一', '工作一時間', '工作一說明', '工作二', '工作二時間', '工作二說明'];
  const widths = [24, 14, 24, 26, 24, 55, 26, 24, 55];
  const sheet = configureExportWorksheet(workbook, '同時段重複工作', headers, widths, [5, 8]);
  const tasks = exportSort(db.tasks);
  const rows = [];
  for (let firstIndex = 0; firstIndex < tasks.length; firstIndex++) {
    const first = tasks[firstIndex];
    const firstPeople = new Set((first.assigned || []).map(String));
    for (let secondIndex = firstIndex + 1; secondIndex < tasks.length; secondIndex++) {
      const second = tasks[secondIndex];
      if (!first.id || first.id === second.id || !overlaps(first, second) || !taskMatchesExportRange(first, range)) continue;
      const sharedPeople = [...new Set((second.assigned || []).map(String))]
        .filter(personId => firstPeople.has(personId))
        .sort((a, b) => personLabel(a).localeCompare(personLabel(b), 'zh-Hant'));
      const sharedDates=taskDates(first).filter(d=>taskIncludesDate(second,d)&&(!range.dateFrom||d>=range.dateFrom)&&(!range.dateTo||d<=range.dateTo));
      for(const overlapDate of sharedDates){const firstWindow=taskWindowOnDate(first,overlapDate),secondWindow=taskWindowOnDate(second,overlapDate),overlapStart=Math.max(firstWindow[0],secondWindow[0]),overlapEnd=Math.min(firstWindow[1],secondWindow[1]),rangeStart=range.timeFrom?exportTimeMinutes(range.timeFrom):0,rangeEnd=range.timeTo?exportTimeMinutes(range.timeTo):1440;if(!(overlapStart<overlapEnd&&overlapStart<rangeEnd&&rangeStart<overlapEnd))continue;for (const personId of sharedPeople) rows.push([
        personLabel(personId), overlapDate, `${exportClockFromSeconds(overlapStart*60)}–${exportClockFromSeconds(overlapEnd*60)}`,
        exportText(first.name), exportTaskTime(first), exportText(first.description),
        exportText(second.name), exportTaskTime(second), exportText(second.description)
      ])}
    }
  }
  rows.sort((a, b) => a[1].localeCompare(b[1]) || a[2].localeCompare(b[2]) || a[0].localeCompare(b[0], 'zh-Hant'));
  appendExportRows(sheet, rows, [5, 8]);
}

function exportUnassignedPeople(workbook, range) {
  const headers = ['服務人員編號', '服務人員', '狀態'];
  const sheet = configureExportWorksheet(workbook, '未安排服務人員', headers, [20, 28, 14]);
  const assignedIds = new Set(db.tasks
    .filter(task => taskMatchesExportRange(task, range))
    .flatMap(task => (task.assigned || []).map(id => String(id))));
  const rows = db.people
    .filter(personRecord => personRecord.active !== false && !assignedIds.has(String(personRecord.id)))
    .sort((a, b) => personLabel(a.id).localeCompare(personLabel(b.id), 'zh-Hant'))
    .map(personRecord => [String(personRecord.id), personRecord.name || '', '啟用中']);
  appendExportRows(sheet, rows);
}

function showExportScope() {
  const field = document.querySelector('#exportScopeField');
  const select = document.querySelector('#exportScope');
  const timeField = document.querySelector('#exportTimeScopeField');
  if (!field || !select) return;
  const needsScope = selectedExportReport === 'personal' || selectedExportReport === 'leaders';
  field.hidden = !needsScope;
  if (timeField) timeField.hidden = selectedExportReport !== 'conflicts' && selectedExportReport !== 'unassigned';
  select.innerHTML = '';
  if (!needsScope) return;
  const allOption = document.createElement('option');
  allOption.value = 'all';
  allOption.textContent = selectedExportReport === 'personal' ? '全部人' : '全部負責人';
  select.append(allOption);
  const choices = selectedExportReport === 'personal'
    ? db.people.map(personRecord => [String(personRecord.id), personRecord.name || personRecord.id])
    : leaderChoices();
  for (const [value, label] of choices) {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = label;
    select.append(option);
  }
}

function openExportDialog() {
  selectedExportReport = '';
  document.querySelectorAll('[data-export-report]').forEach(button => button.classList.remove('selected'));
  document.querySelector('#exportScopeField').hidden = true;
  document.querySelector('#exportTimeScopeField').hidden = true;
  document.querySelectorAll('#exportTimeScopeField input').forEach(input => { input.value = ''; });
  document.querySelector('#downloadExportButton').disabled = true;
  document.querySelector('#exportReportDialog').showModal();
}

async function downloadSelectedReport() {
  if (!selectedExportReport) return;
  if (!window.ExcelJS?.Workbook) {
    toast('Excel 匯出元件尚未載入，請確認網路連線後重試');
    return;
  }
  const selection = document.querySelector('#exportScope').value || 'all';
  const range = exportRangeValues();
  if ((selectedExportReport === 'conflicts' || selectedExportReport === 'unassigned') && !exportRangeIsValid(range)) return;
  const workbook = new ExcelJS.Workbook();
  workbook.creator = '晁人菁英排班系統';
  workbook.created = new Date();
  if (selectedExportReport === 'tasks') exportWorkItems(workbook);
  if (selectedExportReport === 'personal') exportPersonalSchedules(workbook, selection);
  if (selectedExportReport === 'leaders') exportLeaderSheets(workbook, selection);
  if (selectedExportReport === 'conflicts') exportOverlappingAssignments(workbook, range);
  if (selectedExportReport === 'unassigned') exportUnassignedPeople(workbook, range);
  const stamp = exportDateKey();
  const filename = selectedExportReport === 'tasks'
    ? `工作項目_${stamp}.xlsx`
    : selectedExportReport === 'personal'
      ? `個人班表_${selection === 'all' ? '全部人' : safeFilenamePart(person(selection)?.name || selection)}_${stamp}.xlsx`
      : selectedExportReport === 'leaders'
        ? `負責人工作表_${selection === 'all' ? '全部負責人' : safeFilenamePart(leaderChoices().find(([key]) => key === selection)?.[1] || selection)}_${stamp}.xlsx`
        : `${selectedExportReport === 'conflicts' ? '同時段重複工作' : '未安排服務人員'}_${stamp}.xlsx`;
  try {
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    document.querySelector('#exportReportDialog').close();
    toast('Excel 報表已下載');
  } catch (error) {
    console.error('Excel 報表匯出失敗', error);
    toast('Excel 報表匯出失敗，請稍後重試');
  }
}

function safeFilenamePart(value) {
  return String(value || '').replace(/[\\/:*?"<>|]/g, '_').slice(0, 40) || '資料';
}

document.addEventListener('click', event => {
  const button = event.target.closest('[data-action="open-exports"],[data-export-report],[data-action="close-export-dialog"],[data-action="download-export"]');
  if (!button) return;
  if (button.dataset.action === 'open-exports') openExportDialog();
  if (button.dataset.action === 'close-export-dialog') document.querySelector('#exportReportDialog').close();
  if (button.dataset.action === 'download-export') downloadSelectedReport();
  if (button.dataset.exportReport) {
    selectedExportReport = button.dataset.exportReport;
    document.querySelectorAll('[data-export-report]').forEach(item => item.classList.toggle('selected', item === button));
    showExportScope();
    document.querySelector('#downloadExportButton').disabled = false;
  }
});

document.addEventListener('change', event => {
  if (event.target.id === 'exportScope') event.target.dataset.selected = event.target.value;
});

const exportDialogMarkup = `
  <dialog id="exportReportDialog" class="export-report-dialog">
    <div class="modal-head"><div><div class="eyebrow">Excel 報表</div><h2>選擇匯出報表</h2></div><button type="button" class="icon" data-action="close-export-dialog" aria-label="關閉">✕</button></div>
    <div class="export-report-options">
      <button type="button" class="export-report-option" data-export-report="tasks">匯出工作項目</button>
      <button type="button" class="export-report-option" data-export-report="personal">匯出個人班表</button>
      <button type="button" class="export-report-option" data-export-report="leaders">匯出負責人工作表</button>
      <button type="button" class="export-report-option" data-export-report="conflicts">匯出同時段重複工作</button>
      <button type="button" class="export-report-option" data-export-report="unassigned">匯出未安排服務人員</button>
    </div>
    <div class="field" id="exportScopeField" hidden><label for="exportScope">匯出對象</label><select id="exportScope"></select></div>
    <div id="exportTimeScopeField" class="export-time-scope" hidden>
      <label class="field"><span>起始日期</span><input id="exportDateFrom" type="date"></label>
      <label class="field"><span>結束日期</span><input id="exportDateTo" type="date"></label>
      <label class="field"><span>開始時間</span><input id="exportTimeFrom" type="time"></label>
      <label class="field"><span>結束時間</span><input id="exportTimeTo" type="time"></label>
    </div>
    <div class="form-actions"><button type="button" class="btn light" data-action="close-export-dialog">取消</button><button type="button" class="btn primary" id="downloadExportButton" data-action="download-export" disabled>下載 .xlsx</button></div>
  </dialog>`;
document.body.insertAdjacentHTML('beforeend', exportDialogMarkup);
