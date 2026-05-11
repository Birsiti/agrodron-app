// ══════════════════════════════════════════════════════════
// AgroDrone.BY — Google Apps Script
// Версия 1.1 | ЗАО «АТК» | май 2026
// ══════════════════════════════════════════════════════════
// Деплой: Расширения → Apps Script → Деплой → Веб-приложение
// Выполнять как: Я | Доступ: Все
// ══════════════════════════════════════════════════════════

const TELEGRAM_BOT_TOKEN = 'ВАШ_BOT_TOKEN';
const GROUP_CHAT_ID      = 'ВАШ_CHAT_ID';   // Агрообработки 2026

// ── Все листы таблицы ─────────────────────────────────────
const SHEETS = {
  OPERATORY:      'ОПЕРАТОРЫ',
  TEHNIKI:        'ТЕХНИКИ',
  POLETY:         'ПОЛЁТЫ',
  OBRABOTKI:      'ОБРАБОТКИ',
  KLIENTY:        'КЛИЕНТЫ',
  BRIGADY:        'БРИГАДЫ',
  TEHOBSLUZHIVANIE: 'ТЕХНИЧЕСКОЕ ОБСЛУЖИВАНИЕ',
  KOMANDIROVKI:   'КОМАНДИРОВКИ',   // создадим если нет
  DNI:            'ДНИ',            // создадим если нет
  RASHODY:        'РАСХОДЫ'         // создадим если нет
};

// ── Заголовки для новых листов (создаются автоматически) ──
const HEADERS = {
  КОМАНДИРОВКИ: [
    'ID','Дата откр.','Дата закр.','Район','Состав','Техника',
    'Аванс','Потрачено','Остаток','Га итого','Статус','Закрытие','Примечание'
  ],
  ОБРАБОТКИ: [
    'ID','Название','Культура','Препарат','Норма л/га','Га план',
    'Координаты','Агроном ID','Статус','Командировка ID','Га факт','Дата добавления'
  ],
  ДНИ: [
    'ID','Командировка ID','День №','Дата','Га план','Га факт',
    'Полей закрыто','Полей перенесено','Расходы за день','Статус','Примечание'
  ],
  РАСХОДЫ: [
    'ID','Командировка ID','День ID','Дата','Время','Категория','Сумма','Комментарий'
  ]
};

// ══════════════════════════════════════════════════════════
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ══════════════════════════════════════════════════════════

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);

  // Создать лист если не существует
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (HEADERS[name]) {
      sheet.appendRow(HEADERS[name]);
      sheet.getRange(1, 1, 1, HEADERS[name].length)
           .setFontWeight('bold')
           .setBackground('#1b5935')
           .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    Logger.log('Создан лист: ' + name);
  }
  return sheet;
}

function formatDate(date) {
  const d = date.getDate().toString().padStart(2,'0');
  const m = (date.getMonth()+1).toString().padStart(2,'0');
  const y = date.getFullYear();
  return `${d}.${m}.${y}`;
}

function formatTime(date) {
  const h = date.getHours().toString().padStart(2,'0');
  const m = date.getMinutes().toString().padStart(2,'0');
  return `${h}:${m}`;
}

function sheetToObjects(sheet, fromRow) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(fromRow || 1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).filter(r => Object.values(r)[0]); // убираем пустые строки
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return { rowNum: i + 1, rowData: data[i] };
  }
  return null;
}

// ══════════════════════════════════════════════════════════
// ТОЧКА ВХОДА — GET
// ══════════════════════════════════════════════════════════
function doGet(e) {
  // Проверка работоспособности + инициализация листов
  initSheets();
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      service: 'AgroDrone.BY',
      version: '1.1',
      sheets: Object.values(SHEETS)
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Создать все нужные листы при первом запуске
function initSheets() {
  ['КОМАНДИРОВКИ','ОБРАБОТКИ','ДНИ','РАСХОДЫ'].forEach(name => getSheet(name));
}

// ══════════════════════════════════════════════════════════
// ТОЧКА ВХОДА — POST
// ══════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    const data   = body.data || {};
    let result;

    switch (action) {

      // ── ПОЛЯ (ОБРАБОТКИ) ──────────────────────────────
      case 'add_field':           result = addField(data);          break;
      case 'update_field_status': result = updateFieldStatus(data); break;
      case 'get_fields':          result = getFields(data);         break;
      case 'close_field':         result = closeField(data);        break;

      // ── КОМАНДИРОВКИ ──────────────────────────────────
      case 'add_trip':            result = addTrip(data);           break;
      case 'close_trip':          result = closeTrip(data);         break;
      case 'get_trips':           result = getTrips(data);          break;

      // ── ДНИ ───────────────────────────────────────────
      case 'open_day':            result = openDay(data);           break;
      case 'close_day':           result = closeDay(data);          break;

      // ── РАСХОДЫ ───────────────────────────────────────
      case 'add_expense':         result = addExpense(data);        break;
      case 'get_expenses':        result = getExpenses(data);       break;

      // ── ОПЕРАТОРЫ (чтение) ────────────────────────────
      case 'get_operators':       result = getOperators(data);      break;
      case 'get_operator_stats':  result = getOperatorStats(data);  break;

      // ── КЛИЕНТЫ ───────────────────────────────────────
      case 'get_clients':         result = getClients(data);        break;

      // ── ТЕХНИКА ───────────────────────────────────────
      case 'get_equipment':       result = getEquipment(data);      break;
      case 'add_maintenance':     result = addMaintenance(data);    break;

      default:
        result = { error: 'Unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('ERROR: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ══════════════════════════════════════════════════════════
// ПОЛЯ — ОБРАБОТКИ
// ══════════════════════════════════════════════════════════

function addField(data) {
  const sheet = getSheet('ОБРАБОТКИ');
  const id    = 'F-' + Date.now();
  const now   = new Date();

  sheet.appendRow([
    id,
    data.name    || '',
    data.crop    || '',
    data.prep    || '',
    data.norm    || '',
    data.haPlan  || 0,
    data.coords  || '',
    data.agroId  || '',
    'новое',
    '',   // командировка ID
    0,    // га факт
    formatDate(now)
  ]);

  sendTelegram(GROUP_CHAT_ID,
    `🌱 *Новое поле от агронома*\n\n` +
    `📋 ${data.name}\n` +
    `🌾 ${data.crop} · ${data.prep}\n` +
    `📐 ${data.haPlan} га · ${data.norm || '—'} л/га\n` +
    `${data.coords ? '📍 ' + data.coords + '\n' : ''}` +
    `_Требует подтверждения менеджера_`
  );

  return { id, status: 'новое' };
}

function updateFieldStatus(data) {
  const sheet = getSheet('ОБРАБОТКИ');
  const found = findRowById(sheet, data.id);
  if (!found) return { error: 'Поле не найдено: ' + data.id };

  const { rowNum, rowData } = found;
  sheet.getRange(rowNum, 9).setValue(data.status);
  if (data.tripId) sheet.getRange(rowNum, 10).setValue(data.tripId);

  if (data.status === 'подтверждено') {
    sendTelegram(GROUP_CHAT_ID,
      `✅ *Поле подтверждено*\n\n` +
      `📋 ${rowData[1]}\n🌾 ${rowData[2]} · ${rowData[3]}\n` +
      `📐 ${rowData[5]} га\n🚐 Командировка: ${data.tripId || '—'}`
    );
  }

  if (data.status === 'отклонено' && data.agroTelegramId) {
    sendTelegram(data.agroTelegramId,
      `❌ Поле *${rowData[1]}* отклонено.\n${data.reason || ''}`
    );
  }

  return { ok: true };
}

function getFields(data) {
  const sheet  = getSheet('ОБРАБОТКИ');
  const rows   = sheet.getDataRange().getValues();
  if (rows.length < 2) return { fields: [] };

  const fields = rows.slice(1)
    .filter(r => r[0])
    .map(r => ({
      id: r[0], name: r[1], crop: r[2], prep: r[3], norm: r[4],
      haPlan: r[5], coords: r[6], agroId: r[7], status: r[8],
      tripId: r[9], haFact: r[10], dateAdded: r[11]
    }))
    .filter(f => {
      if (data.status && f.status !== data.status) return false;
      if (data.tripId && String(f.tripId) !== String(data.tripId)) return false;
      if (data.agroId && f.agroId !== data.agroId) return false;
      return true;
    });

  return { fields, total: fields.length };
}

function closeField(data) {
  const sheet = getSheet('ОБРАБОТКИ');
  const found = findRowById(sheet, data.id);
  if (!found) return { error: 'Поле не найдено: ' + data.id };

  const { rowNum, rowData } = found;
  const haPlan    = rowData[5];
  const haFactOld = rowData[10] || 0;
  const haFactNew = haFactOld + (data.haToday || 0);
  const remainder = haPlan - haFactNew;
  const status    = (data.isFullyDone || remainder <= 0) ? 'готово' : 'частично';

  sheet.getRange(rowNum, 9).setValue(status);
  sheet.getRange(rowNum, 11).setValue(haFactNew);

  const fieldName = rowData[1];
  const crop      = rowData[2];

  if (status === 'готово') {
    const msg = `✅ *Поле закрыто полностью*\n\n📋 ${fieldName}\n🌾 ${crop}\n📐 ${haFactNew} га из ${haPlan} га`;
    if (data.agroTelegramId) sendTelegram(data.agroTelegramId, msg);
    sendTelegram(GROUP_CHAT_ID, msg);
  } else {
    const msg = `⏩ *Поле частично*\n\n📋 ${fieldName}\n📐 Сегодня: ${data.haToday} га · Всего: ${haFactNew}/${haPlan} га\n🔜 Остаток: *${remainder.toFixed(1)} га*`;
    if (data.agroTelegramId) sendTelegram(data.agroTelegramId, msg);
  }

  return { ok: true, status, haFactNew, remainder: Math.max(0, remainder) };
}

// ══════════════════════════════════════════════════════════
// КОМАНДИРОВКИ
// ══════════════════════════════════════════════════════════

function addTrip(data) {
  const sheet = getSheet('КОМАНДИРОВКИ');
  const id    = data.id || ('TRIP-' + Date.now());
  const now   = formatDate(new Date());

  sheet.appendRow([
    id,
    data.dateStart  || now,
    '',
    data.direction  || '',
    (data.crew  || []).join(', '),
    (data.equip || []).join(', '),
    data.advance    || 0,
    0,              // потрачено
    data.advance    || 0, // остаток
    0,              // га итого
    'активна',
    '',             // закрытие
    ''              // примечание
  ]);

  sendTelegram(GROUP_CHAT_ID,
    `🚐 *Выезд в командировку*\n\n` +
    `📍 ${data.direction}\n` +
    `👥 ${(data.crew||[]).join(', ')}\n` +
    `🚁 ${(data.equip||[]).join(', ')}\n` +
    `💰 Аванс: ${data.advance} BYN · 📅 ${data.dateStart || now}`
  );

  return { id, status: 'активна' };
}

function closeTrip(data) {
  const sheet = getSheet('КОМАНДИРОВКИ');
  const found = findRowById(sheet, data.id);
  if (!found) return { error: 'Командировка не найдена: ' + data.id };

  const { rowNum, rowData } = found;
  const advance  = rowData[6];
  const spent    = data.spent || 0;
  const balance  = advance - spent;
  const dateEnd  = formatDate(new Date());

  sheet.getRange(rowNum, 3).setValue(dateEnd);
  sheet.getRange(rowNum, 8).setValue(spent);
  sheet.getRange(rowNum, 9).setValue(balance);
  sheet.getRange(rowNum, 10).setValue(data.ha || 0);
  sheet.getRange(rowNum, 11).setValue('закрыта');
  sheet.getRange(rowNum, 12).setValue(data.closeOption || '');
  if (data.note) sheet.getRange(rowNum, 13).setValue(data.note);

  const balanceStr = balance >= 0
    ? `Остаток: ${balance.toFixed(2)} BYN (${data.closeOption === 'carry' ? 'перенос' : 'сдано в кассу'})`
    : `Перерасход: ${Math.abs(balance).toFixed(2)} BYN`;

  sendTelegram(GROUP_CHAT_ID,
    `✅ *Командировка закрыта*\n\n` +
    `📍 ${rowData[3]}\n` +
    `🌾 Обработано: *${data.ha || 0} га*\n` +
    `💰 Аванс: ${advance} BYN · Потрачено: ${spent} BYN\n` +
    `${balanceStr}\n` +
    `📅 ${rowData[1]} → ${dateEnd}` +
    `${data.note ? '\n📝 ' + data.note : ''}`
  );

  return { ok: true, balance };
}

function getTrips(data) {
  const sheet = getSheet('КОМАНДИРОВКИ');
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return { trips: [] };

  const trips = rows.slice(1)
    .filter(r => r[0])
    .map(r => ({
      id: r[0], dateStart: r[1], dateEnd: r[2], direction: r[3],
      crew: r[4], equip: r[5], advance: r[6], spent: r[7],
      balance: r[8], ha: r[9], status: r[10]
    }))
    .filter(t => !data.status || t.status === data.status);

  return { trips };
}

// ══════════════════════════════════════════════════════════
// ДНИ
// ══════════════════════════════════════════════════════════

function openDay(data) {
  const sheet    = getSheet('ДНИ');
  const obSheet  = getSheet('ОБРАБОТКИ');
  const id       = 'DAY-' + Date.now();

  sheet.appendRow([
    id,
    data.tripId  || '',
    data.dayNum  || 1,
    data.date    || formatDate(new Date()),
    data.haPlan  || 0,
    0,  // га факт
    0,  // полей закрыто
    0,  // полей перенесено
    0,  // расходы за день
    'активен',
    ''
  ]);

  // Перевести поля в статус 'в работе'
  const fieldIds = (data.fields || []).map(f => String(f.id));
  if (fieldIds.length > 0) {
    const rows = obSheet.getDataRange().getValues();
    rows.forEach((r, i) => {
      if (i === 0) return;
      if (fieldIds.includes(String(r[0])) && r[8] !== 'готово') {
        obSheet.getRange(i + 1, 9).setValue('в работе');
      }
    });
  }

  sendTelegram(GROUP_CHAT_ID,
    `☀️ *День ${data.dayNum} открыт*\n\n` +
    `📍 ${data.tripId}\n` +
    `📅 ${data.date || formatDate(new Date())}\n` +
    `🌾 Полей: ${fieldIds.length} · План: ${data.haPlan} га`
  );

  return { ok: true, dayId: id };
}

function closeDay(data) {
  const sheet   = getSheet('ДНИ');
  const shTrip  = getSheet('КОМАНДИРОВКИ');
  const found   = findRowById(sheet, data.dayId);

  if (found) {
    sheet.getRange(found.rowNum, 6).setValue(data.haDone || 0);
    sheet.getRange(found.rowNum, 7).setValue(data.fieldsDone || 0);
    sheet.getRange(found.rowNum, 8).setValue(data.fieldsCarried || 0);
    sheet.getRange(found.rowNum, 9).setValue(data.spentToday || 0);
    sheet.getRange(found.rowNum, 10).setValue('закрыт');
    if (data.note) sheet.getRange(found.rowNum, 11).setValue(data.note);
  }

  // Обновить командировку
  if (data.tripId) {
    const tripFound = findRowById(shTrip, data.tripId);
    if (tripFound) {
      const advance = tripFound.rowData[6];
      shTrip.getRange(tripFound.rowNum, 8).setValue(data.totalSpent || 0);
      shTrip.getRange(tripFound.rowNum, 9).setValue(advance - (data.totalSpent || 0));
      shTrip.getRange(tripFound.rowNum, 10).setValue(data.totalHa || 0);
    }
  }

  const carriedStr = data.fieldsCarried > 0 ? `\n⏩ Перенесено: ${data.fieldsCarried} пол.` : '';

  sendTelegram(GROUP_CHAT_ID,
    `🌙 *День ${data.dayNum} закрыт*\n\n` +
    `📅 ${data.date}\n` +
    `🌾 Обработано: *${data.haDone} га*\n` +
    `✅ Полей закрыто: ${data.fieldsDone}${carriedStr}\n` +
    `💰 Расходы за день: ${data.spentToday || 0} BYN` +
    `${data.note ? '\n📝 ' + data.note : ''}`
  );

  return { ok: true };
}

// ══════════════════════════════════════════════════════════
// РАСХОДЫ
// ══════════════════════════════════════════════════════════

function addExpense(data) {
  const sheet = getSheet('РАСХОДЫ');
  const now   = new Date();
  const id    = 'EXP-' + Date.now();

  sheet.appendRow([
    id,
    data.tripId   || '',
    data.dayId    || '',
    formatDate(now),
    formatTime(now),
    data.label    || data.cat || '',
    data.amount   || 0,
    data.comment  || ''
  ]);

  // Обновить потраченное в командировке
  if (data.tripId && data.totalSpent !== undefined) {
    const shTrip = getSheet('КОМАНДИРОВКИ');
    const found  = findRowById(shTrip, data.tripId);
    if (found) {
      const advance = found.rowData[6];
      shTrip.getRange(found.rowNum, 8).setValue(data.totalSpent);
      shTrip.getRange(found.rowNum, 9).setValue(advance - data.totalSpent);
    }
  }

  return { id, ok: true };
}

function getExpenses(data) {
  const sheet = getSheet('РАСХОДЫ');
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return { expenses: [] };

  const expenses = rows.slice(1)
    .filter(r => r[0])
    .map(r => ({
      id: r[0], tripId: r[1], dayId: r[2], date: r[3],
      time: r[4], label: r[5], amount: r[6], comment: r[7]
    }))
    .filter(e => !data.tripId || String(e.tripId) === String(data.tripId));

  return { expenses };
}

// ══════════════════════════════════════════════════════════
// ОПЕРАТОРЫ (только чтение)
// ══════════════════════════════════════════════════════════

function getOperators(data) {
  const sheet = getSheet('ОПЕРАТОРЫ');
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return { operators: [] };

  const operators = rows.slice(1)
    .filter(r => r[0])
    .map(r => ({
      id: r[0], name: r[1], phone: r[2], region: r[3],
      cert: r[4], hours: r[5], status: r[6], regDate: r[7],
      telegramId: r[8], type: r[9], hectares: r[10], trips: r[11]
    }))
    .filter(o => !data.status || o.status === data.status);

  return { operators };
}

function getOperatorStats(data) {
  const sheet = getSheet('ОПЕРАТОРЫ');
  const found = findRowById(sheet, data.id);
  if (!found) return { error: 'Оператор не найден' };
  const r = found.rowData;
  return {
    id: r[0], name: r[1], hours: r[5],
    status: r[6], hectares: r[10], trips: r[11]
  };
}

// ══════════════════════════════════════════════════════════
// КЛИЕНТЫ (только чтение)
// ══════════════════════════════════════════════════════════

function getClients(data) {
  const sheet = getSheet('КЛИЕНТЫ');
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return { clients: [] };

  return {
    clients: rows.slice(1).filter(r => r[0]).map(r => ({
      id: r[0], name: r[1], phone: r[2], region: r[3], contact: r[4]
    }))
  };
}

// ══════════════════════════════════════════════════════════
// ТЕХНИКА
// ══════════════════════════════════════════════════════════

function getEquipment(data) {
  const sheet = getSheet('ТЕХНИКИ');
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return { equipment: [] };

  return {
    equipment: rows.slice(1).filter(r => r[0]).map(r => ({
      id: r[0], name: r[1], model: r[2], status: r[3], hours: r[4]
    }))
  };
}

function addMaintenance(data) {
  const sheet = getSheet('ТЕХНИЧЕСКОЕ ОБСЛУЖИВАНИЕ');
  const id    = 'TO-' + Date.now();

  sheet.appendRow([
    id,
    data.equipId  || '',
    formatDate(new Date()),
    data.type     || '',
    data.desc     || '',
    data.operatorId || '',
    data.cost     || 0
  ]);

  return { id, ok: true };
}

// ══════════════════════════════════════════════════════════
// TELEGRAM
// ══════════════════════════════════════════════════════════

function sendTelegram(chatId, text) {
  if (!chatId || String(chatId).startsWith('ВАШ')) return;
  try {
    UrlFetchApp.fetch(
      `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          chat_id: chatId,
          text: text,
          parse_mode: 'Markdown'
        })
      }
    );
  } catch(e) {
    Logger.log('Telegram error: ' + e.message);
  }
}
