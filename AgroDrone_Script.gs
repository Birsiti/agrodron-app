// ══════════════════════════════════════════════════════════
// AgroDrone.BY — Google Apps Script
// Версия 1.0 | ЗАО «АТК» | май 2026
// ══════════════════════════════════════════════════════════
// Деплой: Расширения → Apps Script → Деплой → Веб-приложение
// Доступ: Все (анонимный)
// ══════════════════════════════════════════════════════════

const TELEGRAM_BOT_TOKEN = 'ВАШ_BOT_TOKEN'; // @AgroDrone_BY_Bot
const GROUP_CHAT_ID      = 'ВАШ_CHAT_ID';   // Агрообработки 2026
// Агроном — личный Chat ID (для уведомлений по полям)
// Передаётся в запросе из интерфейса

// Названия листов
const SHEET_OBRABOTKI   = 'ОБРАБОТКИ';
const SHEET_KOMANDIROVKI = 'КОМАНДИРОВКИ';
const SHEET_POLETY       = 'ПОЛЁТЫ';

// ══════════════════════════════════════════════════════════
// ТОЧКА ВХОДА — GET (для проверки что скрипт живой)
// ══════════════════════════════════════════════════════════
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', service: 'AgroDrone.BY', version: '1.0' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════
// ТОЧКА ВХОДА — POST (все действия)
// ══════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    const data   = body.data || {};

    let result;

    switch (action) {

      // ── ПОЛЯ ──────────────────────────────────────────
      case 'add_field':
        result = addField(data);
        break;

      case 'update_field_status':
        result = updateFieldStatus(data);
        break;

      case 'get_fields':
        result = getFields(data);
        break;

      case 'close_field':
        result = closeField(data);
        break;

      // ── КОМАНДИРОВКИ ──────────────────────────────────
      case 'add_trip':
        result = addTrip(data);
        break;

      case 'close_trip':
        result = closeTrip(data);
        break;

      case 'get_trips':
        result = getTrips(data);
        break;

      // ── ДНИ ───────────────────────────────────────────
      case 'open_day':
        result = openDay(data);
        break;

      case 'close_day':
        result = closeDay(data);
        break;

      default:
        result = { error: 'Unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ══════════════════════════════════════════════════════════
// ПОЛЯ — ОБРАБОТКИ
// Структура: A=ID B=Название C=Культура D=Препарат E=Норма
//            F=Га план G=Координаты H=Агроном ID I=Статус
//            J=Командировка ID K=Га факт L=Дата добавления
// ══════════════════════════════════════════════════════════

function addField(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_OBRABOTKI);

  const id   = 'F-' + Date.now();
  const date = formatDate(new Date());

  const row = [
    id,
    data.name     || '',
    data.crop     || '',
    data.prep     || '',
    data.norm     || '',
    data.haPlan   || 0,
    data.coords   || '',
    data.agroId   || '',
    'новое',        // статус
    '',             // командировка ID
    0,              // га факт
    date
  ];

  sheet.appendRow(row);

  // Уведомление менеджеру
  const msg = `🌱 *Новое поле от агронома*\n\n` +
    `📋 ${data.name}\n` +
    `🌾 ${data.crop} · ${data.prep}\n` +
    `📐 ${data.haPlan} га · ${data.norm} л/га\n` +
    `${data.coords ? '📍 ' + data.coords + '\n' : ''}` +
    `\n_Требует подтверждения менеджера_`;

  sendTelegram(GROUP_CHAT_ID, msg);

  return { id, status: 'новое' };
}

// ──────────────────────────────────────────────────────────
function updateFieldStatus(data) {
  // data: { id, status, tripId?, agroTelegramId? }
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_OBRABOTKI);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.getRange(i + 1, 9).setValue(data.status);             // I = статус
      if (data.tripId) sheet.getRange(i + 1, 10).setValue(data.tripId); // J = командировка

      // Уведомления
      if (data.status === 'подтверждено') {
        const msg = `✅ *Поле подтверждено менеджером*\n\n` +
          `📋 ${rows[i][1]}\n` +
          `🌾 ${rows[i][2]} · ${rows[i][3]}\n` +
          `📐 ${rows[i][5]} га\n` +
          `🚐 Командировка: ${data.tripId || '—'}`;
        sendTelegram(GROUP_CHAT_ID, msg);
      }

      if (data.status === 'отклонено' && data.agroTelegramId) {
        const msg = `❌ Поле *${rows[i][1]}* отклонено менеджером.\n${data.reason || ''}`;
        sendTelegram(data.agroTelegramId, msg);
      }

      return { ok: true, row: i + 1 };
    }
  }
  return { error: 'Field not found: ' + data.id };
}

// ──────────────────────────────────────────────────────────
function getFields(data) {
  // data: { status?, tripId?, agroId? }
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_OBRABOTKI);
  const rows  = sheet.getDataRange().getValues();

  const fields = rows.slice(1)
    .filter(r => r[0]) // есть ID
    .map(r => ({
      id:        r[0],
      name:      r[1],
      crop:      r[2],
      prep:      r[3],
      norm:      r[4],
      haPlan:    r[5],
      coords:    r[6],
      agroId:    r[7],
      status:    r[8],
      tripId:    r[9],
      haFact:    r[10],
      dateAdded: r[11]
    }))
    .filter(f => {
      if (data.status && f.status !== data.status) return false;
      if (data.tripId && f.tripId !== data.tripId)  return false;
      if (data.agroId && f.agroId !== data.agroId)  return false;
      return true;
    });

  return { fields };
}

// ──────────────────────────────────────────────────────────
function closeField(data) {
  // data: { id, haToday, isFullyDone, agroTelegramId }
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_OBRABOTKI);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      const haPlan    = rows[i][5];
      const haFactOld = rows[i][10] || 0;
      const haFactNew = haFactOld + (data.haToday || 0);
      const remainder = haPlan - haFactNew;
      const status    = data.isFullyDone || remainder <= 0 ? 'готово' : 'частично';

      sheet.getRange(i + 1, 9).setValue(status);       // I = статус
      sheet.getRange(i + 1, 11).setValue(haFactNew);   // K = га факт

      // Уведомление агроному
      const fieldName = rows[i][1];
      if (status === 'готово') {
        const msg = `✅ *Поле закрыто полностью*\n\n` +
          `📋 ${fieldName}\n` +
          `🌾 ${rows[i][2]}\n` +
          `📐 Обработано: *${haFactNew} га* из ${haPlan} га`;
        if (data.agroTelegramId) sendTelegram(data.agroTelegramId, msg);
        sendTelegram(GROUP_CHAT_ID, msg);
      } else {
        const msg = `⏩ *Поле обработано частично*\n\n` +
          `📋 ${fieldName}\n` +
          `📐 Сегодня: ${data.haToday} га\n` +
          `✅ Всего: ${haFactNew} га из ${haPlan} га\n` +
          `🔜 Остаток: *${remainder.toFixed(1)} га* — перенесено`;
        if (data.agroTelegramId) sendTelegram(data.agroTelegramId, msg);
      }

      return { ok: true, status, haFactNew, remainder: Math.max(0, remainder) };
    }
  }
  return { error: 'Field not found: ' + data.id };
}

// ══════════════════════════════════════════════════════════
// КОМАНДИРОВКИ
// Структура: A=ID B=Дата откр. C=Дата закр. D=Район
//            E=Состав F=Техника G=Аванс H=Потрачено
//            I=Остаток J=Га итого K=Статус L=Закрытие(сдал/перенос)
// ══════════════════════════════════════════════════════════

function addTrip(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_KOMANDIROVKI);

  // Создать лист если нет
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_KOMANDIROVKI);
    sheet.appendRow(['ID','Дата откр.','Дата закр.','Район','Состав','Техника',
                     'Аванс','Потрачено','Остаток','Га итого','Статус','Закрытие']);
  }

  const id  = data.id || ('TRIP-' + Date.now());
  const row = [
    id,
    data.dateStart  || formatDate(new Date()),
    '',
    data.direction  || '',
    (data.crew   || []).join(', '),
    (data.equip  || []).join(', '),
    data.advance    || 0,
    0,              // потрачено
    data.advance    || 0, // остаток
    0,              // га итого
    'активна',
    ''
  ];

  sheet.appendRow(row);

  // Уведомление в группу
  const msg = `🚐 *Выезд в командировку*\n\n` +
    `📍 ${data.direction}\n` +
    `👥 Состав: ${(data.crew || []).join(', ')}\n` +
    `🚁 Техника: ${(data.equip || []).join(', ')}\n` +
    `💰 Аванс: ${data.advance} BYN\n` +
    `📅 ${data.dateStart}`;

  sendTelegram(GROUP_CHAT_ID, msg);

  return { id, status: 'активна' };
}

// ──────────────────────────────────────────────────────────
function closeTrip(data) {
  // data: { id, spent, ha, closeOption, note }
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_KOMANDIROVKI);
  if (!sheet) return { error: 'Sheet not found' };

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      const advance  = rows[i][6];
      const spent    = data.spent || 0;
      const balance  = advance - spent;
      const dateEnd  = formatDate(new Date());

      sheet.getRange(i + 1, 3).setValue(dateEnd);
      sheet.getRange(i + 1, 8).setValue(spent);
      sheet.getRange(i + 1, 9).setValue(balance);
      sheet.getRange(i + 1, 10).setValue(data.ha || 0);
      sheet.getRange(i + 1, 11).setValue('закрыта');
      sheet.getRange(i + 1, 12).setValue(data.closeOption || '');

      const balanceStr = balance >= 0
        ? `Остаток: ${balance.toFixed(2)} BYN (${data.closeOption === 'carry' ? 'перенос' : 'сдано в кассу'})`
        : `Перерасход: ${Math.abs(balance).toFixed(2)} BYN`;

      const msg = `✅ *Командировка закрыта*\n\n` +
        `📍 ${rows[i][3]}\n` +
        `🌾 Га обработано: *${data.ha || 0} га*\n` +
        `💰 Аванс: ${advance} BYN · Потрачено: ${spent} BYN\n` +
        `${balanceStr}\n` +
        `📅 ${rows[i][1]} → ${dateEnd}` +
        `${data.note ? '\n📝 ' + data.note : ''}`;

      sendTelegram(GROUP_CHAT_ID, msg);

      return { ok: true, balance };
    }
  }
  return { error: 'Trip not found: ' + data.id };
}

// ──────────────────────────────────────────────────────────
function getTrips(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_KOMANDIROVKI);
  if (!sheet) return { trips: [] };

  const rows = sheet.getDataRange().getValues();
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
  // data: { tripId, dayNum, date, fields[], haPlan }
  // Меняем статус полей на 'в работе'
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_OBRABOTKI);
  const rows  = sheet.getDataRange().getValues();

  const fieldIds = (data.fields || []).map(f => f.id);
  fieldIds.forEach(fid => {
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === fid && rows[i][8] !== 'готово') {
        sheet.getRange(i + 1, 9).setValue('в работе');
      }
    }
  });

  // Уведомление в группу
  const msg = `☀️ *День ${data.dayNum} открыт*\n\n` +
    `📍 Командировка: ${data.tripId}\n` +
    `📅 ${data.date}\n` +
    `🌾 Полей: ${fieldIds.length} · План: ${data.haPlan} га`;

  sendTelegram(GROUP_CHAT_ID, msg);

  return { ok: true };
}

// ──────────────────────────────────────────────────────────
function closeDay(data) {
  // data: { tripId, dayNum, date, haDone, fieldsDone, fieldsCarried, spent, note }

  // Обновить потраченное в командировке
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const shTrip  = ss.getSheetByName(SHEET_KOMANDIROVKI);
  if (shTrip) {
    const rows = shTrip.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.tripId) {
        const advance = rows[i][6];
        shTrip.getRange(i + 1, 8).setValue(data.totalSpent || 0);
        shTrip.getRange(i + 1, 9).setValue(advance - (data.totalSpent || 0));
        shTrip.getRange(i + 1, 10).setValue(data.totalHa || 0);
        break;
      }
    }
  }

  // Уведомление в группу
  const carriedStr = data.fieldsCarried > 0
    ? `\n⏩ Перенесено полей: ${data.fieldsCarried}` : '';

  const msg = `🌙 *День ${data.dayNum} закрыт*\n\n` +
    `📅 ${data.date}\n` +
    `🌾 Га обработано: *${data.haDone} га*\n` +
    `✅ Полей закрыто: ${data.fieldsDone}` +
    `${carriedStr}\n` +
    `💰 Расходы за день: ${data.spentToday || 0} BYN` +
    `${data.note ? '\n📝 ' + data.note : ''}`;

  sendTelegram(GROUP_CHAT_ID, msg);

  return { ok: true };
}

// ══════════════════════════════════════════════════════════
// TELEGRAM
// ══════════════════════════════════════════════════════════

function sendTelegram(chatId, text) {
  if (!chatId || chatId.startsWith('ВАШ')) return; // не отправлять если не настроено
  try {
    const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: chatId,
        text: text,
        parse_mode: 'Markdown'
      })
    });
  } catch(e) {
    console.log('Telegram error:', e.message);
  }
}

// ══════════════════════════════════════════════════════════
// UTILS
// ══════════════════════════════════════════════════════════

function formatDate(date) {
  const d = date.getDate().toString().padStart(2,'0');
  const m = (date.getMonth()+1).toString().padStart(2,'0');
  const y = date.getFullYear();
  return `${d}.${m}.${y}`;
}
