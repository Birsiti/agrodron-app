// ══════════════════════════════════════════════════════════
// AgroDrone.BY — Google Apps Script v2.2
// ЗАО «АТК» | май 2026
// Изменения v2.2: добавлена функция updateField (редактирование поля)
// ══════════════════════════════════════════════════════════

const TELEGRAM_BOT_TOKEN = '8675713209:AAHS6P8TGpO3efMHwDblTzyLwDarHMINiAk';
const GROUP_CHAT_ID      = '-1003761125572';

// ══════════════════════════════════════════════════════════
// СТРУКТУРА ЛИСТОВ
// ══════════════════════════════════════════════════════════
//
// ОПЕРАТОРЫ: A=ID B=ФИО C=Телефон D=Район E=Серт.БАК
//   F=Часов G=Статус H=Дата рег. I=Telegram ID J=Тип
//   K=Га за сезон L=Выездов M=Роль N=Дата посл.выезда
//
// КОМАНДИРОВКИ: A=ID B=Дата откр. C=Дата закр. D=Район
//   E=Состав F=Техника G=Аванс H=Потрачено I=Остаток
//   J=Га итого K=Статус L=Закрытие M=Примечание
//
// ДНИ: A=ID B=Командировка ID C=День№ D=Дата E=Га план
//   F=Га факт G=Полей закрыто H=Полей перенесено
//   I=Расходы за день J=Статус K=Примечание
//
// РАСХОДЫ: A=ID B=Командировка ID C=День ID D=Дата
//   E=Время F=Категория G=Сумма H=Комментарий
//
// ПОЛЯ КЛИЕНТОВ: A=ID B=ID клиента C=Районный центр
//   D=Название поля E=Ссылка Maps F=Культура G=Препарат
//   H=Норма л/га I=Площадь га J=Статус K=Фото URL
//   L=Дата закрытия M=ID обработки N=Примечания
//
// ОБРАБОТКИ: A=ID B=Дата C=ID заявки D=НП E=Культура
//   F=Препарат G=Норма л/га H=Га факт I=Борт
//   J=Операторы K=Га/оператор L=Фото URL M=Статус
//   N=Дата закрытия O=Примечания
//
// КЛИЕНТЫ: A=ID B=Хозяйство C=Агроном ФИО D=Telegram ID
//   E=Телефон F=Email G=Статус H=Дата активации
//   I=Менеджер J=Примечания
//
// БРИГАДЫ: A=ID B=Дата создания C=Старший(TG) D=Операторы(TG)
//   E=Техник(TG) F=ID выезда G=Статус H=Дата завершения I=Примечания
//
// ЗАЯВКИ: A=ID B=Дата C=Организация D=УНП E=Телефон
//   F=Telegram ID G=Район/НП H=Площадь I=Культура J=Коэф.культуры
//   K=Сложность L=Препятствия M=Участков N=Срочность O=Сроки
//   P=Комментарий Q=Статус R=Дата статуса S=Итог.коэф.
//   T=Обработка BYN U=Дорога BYN V=Итого с НДС W=Цена за га
//   X=Заметки менеджера Y=Расстояние км
//
// КАНДИДАТЫ: A=ID B=Дата C=ФИО D=Возраст E=Район/город
//   F=Телефон G=Роль H=Опыт I=Авто J=Выезды K=Источник
//   L=О себе M=Telegram ID N=Статус O=Дата статуса P=Заметки
// ══════════════════════════════════════════════════════════

// ══════════════════════════════════════════════════════════
// UTILS
// ══════════════════════════════════════════════════════════

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Лист не найден: ' + name);
  return sheet;
}

function formatDate(date) {
  const d = date.getDate().toString().padStart(2,'0');
  const m = (date.getMonth()+1).toString().padStart(2,'0');
  const y = date.getFullYear();
  return d + '.' + m + '.' + y;
}

function formatTime(date) {
  return date.getHours().toString().padStart(2,'0') + ':' +
         date.getMinutes().toString().padStart(2,'0');
}

function normalizeNumber(val) {
  if (val === null || val === undefined || val === '') return '';
  return String(val).replace(',', '.');
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      return { rowNum: i + 1, rowData: data[i] };
    }
  }
  return null;
}

function sheetRows(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).filter(r => r[0] !== '' && r[0] !== null);
}

// ══════════════════════════════════════════════════════════
// ТОЧКА ВХОДА — GET
// ══════════════════════════════════════════════════════════

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok', service: 'AgroDrone.BY', version: '2.2',
      time: formatDate(new Date())
    }))
    .setMimeType(ContentService.MimeType.JSON);
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
      case 'add_field':             result = addField(data);            break;
      case 'get_fields':            result = getFields(data);           break;
      case 'update_field':          result = updateField(data);         break;  // v2.2
      case 'update_field_status':   result = updateFieldStatus(data);   break;
      case 'close_field':           result = closeField(data);          break;
      case 'get_clients':           result = getClients(data);          break;
      case 'get_client':            result = getClient(data);           break;
      case 'add_trip':              result = addTrip(data);             break;
      case 'get_trips':             result = getTrips(data);            break;
      case 'close_trip':            result = closeTrip(data);           break;
      case 'open_day':              result = openDay(data);             break;
      case 'close_day':             result = closeDay(data);            break;
      case 'add_expense':           result = addExpense(data);          break;
      case 'get_expenses':          result = getExpenses(data);         break;
      case 'add_obrabotka':         result = addObrabotka(data);        break;
      case 'get_obrabotki':         result = getObrabotki(data);        break;
      case 'get_operators':         result = getOperators(data);        break;
      case 'update_operator_stats': result = updateOperatorStats(data); break;
      case 'get_zayavki':           result = getZayavki(data);          break;
      case 'update_zayavka':        result = updateZayavka(data);       break;
      case 'get_kandidaty':         result = getKandidaty(data);        break;
      case 'update_kandidat':       result = updateKandidat(data);      break;
      case 'get_brigady':           result = getBrigady(data);          break;
      default: result = { error: 'Unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('ERROR: ' + err.message + '\n' + err.stack);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ══════════════════════════════════════════════════════════
// ПОЛЯ КЛИЕНТОВ
// A=ID B=ID клиента C=Районный центр D=Название поля
// E=Ссылка Maps F=Культура G=Препарат H=Норма л/га
// I=Площадь га J=Статус K=Фото URL L=Дата закрытия
// M=ID обработки N=Примечания
// ══════════════════════════════════════════════════════════

function addField(data) {
  const sheet = getSheet('ПОЛЯ КЛИЕНТОВ');
  const id    = 'PK-' + Date.now();

  sheet.appendRow([
    id,                                                // A=ID
    data.clientId || '',                               // B=ID клиента
    data.city     || '',                               // C=Районный центр
    data.name     || '',                               // D=Название поля
    data.mapsUrl  || data.coords || '',                // E=Ссылка Maps
    data.crop     || '',                               // F=Культура
    data.prep     || '',                               // G=Препарат
    normalizeNumber(data.norm) || '',                  // H=Норма л/га
    parseFloat(normalizeNumber(data.haPlan)) || 0,     // I=Площадь га
    'новое',                                           // J=Статус
    '',                                                // K=Фото URL
    '',                                                // L=Дата закрытия
    '',                                                // M=ID обработки
    data.note || ''                                    // N=Примечания
  ]);

  sendTelegram(GROUP_CHAT_ID,
    `🌱 *Новое поле от агронома*\n\n` +
    `📍 ${data.city || '—'}\n` +
    `📋 ${data.name}\n` +
    `🌾 ${data.crop} · ${data.prep}\n` +
    `📐 ${data.haPlan} га · ${data.norm || '—'} л/га\n` +
    `${data.mapsUrl || data.coords ? '🗺 ' + (data.mapsUrl || data.coords) + '\n' : ''}` +
    `_Требует подтверждения менеджера_`
  );

  return { id, status: 'новое' };
}

function getFields(data) {
  const sheet = getSheet('ПОЛЯ КЛИЕНТОВ');
  const rows  = sheetRows(sheet);

  const fields = rows.map(r => ({
    id:          String(r[0]),
    clientId:    String(r[1]),
    city:        r[2],
    name:        r[3],
    mapsUrl:     r[4],
    crop:        r[5],
    prep:        r[6],
    norm:        r[7],
    haPlan:      r[8],
    status:      r[9] || 'новое',
    photoUrl:    r[10],
    dateClosed:  r[11],
    obrabotkaId: r[12],
    note:        r[13]
  })).filter(f => {
    if (data.clientId && f.clientId !== String(data.clientId)) return false;
    if (data.status   && f.status   !== data.status)           return false;
    if (data.city     && f.city     !== data.city)             return false;
    return true;
  });

  return { fields, total: fields.length };
}

// ── НОВАЯ ФУНКЦИЯ v2.2 ─────────────────────────────────────
function updateField(data) {
  const sheet = getSheet('ПОЛЯ КЛИЕНТОВ');
  const found = findRowById(sheet, data.id);
  if (!found) return { error: 'Поле не найдено: ' + data.id };

  const { rowNum } = found;

  if (data.name    !== undefined) sheet.getRange(rowNum, 4).setValue(data.name);    // D=Название
  if (data.mapsUrl !== undefined) sheet.getRange(rowNum, 5).setValue(data.mapsUrl); // E=Ссылка Maps
  if (data.crop    !== undefined) sheet.getRange(rowNum, 6).setValue(data.crop);    // F=Культура
  if (data.prep    !== undefined) sheet.getRange(rowNum, 7).setValue(data.prep);    // G=Препарат
  if (data.norm    !== undefined) sheet.getRange(rowNum, 8).setValue(data.norm);    // H=Норма
  if (data.haPlan  !== undefined) sheet.getRange(rowNum, 9).setValue(data.haPlan);  // I=Площадь

  return { ok: true, id: data.id };
}
// ───────────────────────────────────────────────────────────

function updateFieldStatus(data) {
  const sheet = getSheet('ПОЛЯ КЛИЕНТОВ');
  const found = findRowById(sheet, data.id);
  if (!found) return { error: 'Поле не найдено: ' + data.id };

  const { rowNum, rowData } = found;
  sheet.getRange(rowNum, 10).setValue(data.status); // J=Статус

  if (data.status === 'подтверждено') {
    sendTelegram(GROUP_CHAT_ID,
      `✅ *Поле подтверждено менеджером*\n\n` +
      `📍 ${rowData[2]}\n📋 ${rowData[3]}\n` +
      `🌾 ${rowData[5]} · ${rowData[6]}\n` +
      `📐 ${rowData[8]} га`
    );
  }
  if (data.status === 'отклонено' && data.agroTelegramId) {
    sendTelegram(data.agroTelegramId,
      `❌ Поле *${rowData[3]}* отклонено.\n${data.reason || ''}`
    );
  }

  return { ok: true, status: data.status };
}

function closeField(data) {
  const sheet = getSheet('ПОЛЯ КЛИЕНТОВ');
  const found = findRowById(sheet, data.id);
  if (!found) return { error: 'Поле не найдено: ' + data.id };

  const { rowNum, rowData } = found;
  const haPlan    = parseFloat(rowData[8]) || 0;  // I=Площадь га
  const haToday   = parseFloat(data.haToday) || 0;
  const remainder = haPlan - haToday;
  const isDone    = data.isFullyDone || remainder <= 0;
  const status    = isDone ? 'готово' : 'частично';
  const now       = formatDate(new Date());

  sheet.getRange(rowNum, 10).setValue(status);           // J=Статус
  if (isDone) sheet.getRange(rowNum, 12).setValue(now);  // L=Дата закрытия

  const obrabotkaId = addObrabotka({
    fieldName: rowData[3],
    crop:      rowData[5],
    prep:      rowData[6],
    norm:      rowData[7],
    haFact:    haToday,
    bort:      data.bort || '',
    operators: (data.operatorIds || []).join(', '),
    photoUrl:  data.photoUrl  || '',
    zayavkaId: data.zayavkaId || '',
    note:      data.note      || ''
  }).id;

  sheet.getRange(rowNum, 13).setValue(obrabotkaId); // M=ID обработки

  const fieldName = rowData[3];
  const crop      = rowData[5];

  if (isDone) {
    const msg = `✅ *Поле закрыто полностью*\n\n📋 ${fieldName}\n🌾 ${crop}\n📐 ${haToday} га из ${haPlan} га`;
    if (data.agroTelegramId) sendTelegram(data.agroTelegramId, msg);
    sendTelegram(GROUP_CHAT_ID, msg);
  } else {
    const msg = `⏩ *Поле частично*\n\n📋 ${fieldName}\n📐 Сегодня: ${haToday} га\n🔜 Остаток: *${remainder.toFixed(1)} га*`;
    if (data.agroTelegramId) sendTelegram(data.agroTelegramId, msg);
  }

  return { ok: true, status, obrabotkaId, remainder: Math.max(0, remainder) };
}

// ══════════════════════════════════════════════════════════
// КЛИЕНТЫ
// ══════════════════════════════════════════════════════════

function getClients(data) {
  const sheet = getSheet('КЛИЕНТЫ');
  const rows  = sheetRows(sheet);
  const clients = rows.map(r => ({
    id: String(r[0]), name: r[1], agroName: r[2],
    telegramId: String(r[3]), phone: r[4], email: r[5],
    status: r[6], dateActive: r[7], manager: r[8], note: r[9]
  })).filter(c => !data.status || c.status === data.status);
  return { clients, total: clients.length };
}

function getClient(data) {
  const sheet = getSheet('КЛИЕНТЫ');
  const found = findRowById(sheet, data.id);
  if (!found) return { error: 'Клиент не найден' };
  const r = found.rowData;
  return { id: String(r[0]), name: r[1], agroName: r[2], telegramId: String(r[3]), phone: r[4], status: r[6] };
}

// ══════════════════════════════════════════════════════════
// КОМАНДИРОВКИ
// ══════════════════════════════════════════════════════════

function addTrip(data) {
  const sheet = getSheet('КОМАНДИРОВКИ');
  const id    = data.id || ('TRIP-' + Date.now());
  const now   = formatDate(new Date());
  const adv   = parseFloat(data.advance) || 0;

  sheet.appendRow([
    id, data.dateStart || now, '', data.direction || '',
    (data.crew  || []).join(', '), (data.equip || []).join(', '),
    adv, 0, adv, 0, 'активна', '', ''
  ]);

  sendTelegram(GROUP_CHAT_ID,
    `🚐 *Выезд в командировку*\n\n` +
    `📍 ${data.direction}\n` +
    `👥 ${(data.crew||[]).join(', ')}\n` +
    `🚁 ${(data.equip||[]).join(', ')}\n` +
    `💰 Аванс: ${adv} BYN · 📅 ${data.dateStart || now}`
  );

  return { id, status: 'активна' };
}

function getTrips(data) {
  const sheet = getSheet('КОМАНДИРОВКИ');
  const rows  = sheetRows(sheet);
  const trips = rows.map(r => ({
    id: String(r[0]), dateStart: r[1], dateEnd: r[2], direction: r[3],
    crew: r[4], equip: r[5], advance: r[6], spent: r[7], balance: r[8],
    ha: r[9], status: r[10], closeOption: r[11], note: r[12]
  })).filter(t => !data.status || t.status === data.status);
  return { trips, total: trips.length };
}

function closeTrip(data) {
  const sheet = getSheet('КОМАНДИРОВКИ');
  const found = findRowById(sheet, data.id);
  if (!found) return { error: 'Командировка не найдена: ' + data.id };

  const { rowNum, rowData } = found;
  const advance = parseFloat(rowData[6]) || 0;
  const spent   = parseFloat(data.spent) || 0;
  const balance = advance - spent;
  const now     = formatDate(new Date());

  sheet.getRange(rowNum, 3).setValue(now);
  sheet.getRange(rowNum, 8).setValue(spent);
  sheet.getRange(rowNum, 9).setValue(balance);
  sheet.getRange(rowNum, 10).setValue(data.ha || 0);
  sheet.getRange(rowNum, 11).setValue('закрыта');
  sheet.getRange(rowNum, 12).setValue(data.closeOption || '');
  if (data.note) sheet.getRange(rowNum, 13).setValue(data.note);

  const balStr = balance >= 0
    ? `Остаток: ${balance.toFixed(2)} BYN (${data.closeOption === 'carry' ? 'перенос' : 'сдано в кассу'})`
    : `Перерасход: ${Math.abs(balance).toFixed(2)} BYN`;

  sendTelegram(GROUP_CHAT_ID,
    `✅ *Командировка закрыта*\n\n📍 ${rowData[3]}\n` +
    `🌾 Обработано: *${data.ha || 0} га*\n` +
    `💰 Аванс: ${advance} · Потрачено: ${spent} BYN\n${balStr}\n` +
    `📅 ${rowData[1]} → ${now}${data.note ? '\n📝 ' + data.note : ''}`
  );

  return { ok: true, balance };
}

// ══════════════════════════════════════════════════════════
// ДНИ
// ══════════════════════════════════════════════════════════

function openDay(data) {
  const sheet = getSheet('ДНИ');
  const id    = 'DAY-' + Date.now();
  const now   = data.date || formatDate(new Date());

  sheet.appendRow([id, data.tripId||'', data.dayNum||1, now, data.haPlan||0, 0, 0, 0, 0, 'активен', '']);

  if (data.fieldIds && data.fieldIds.length > 0) {
    const pkSheet = getSheet('ПОЛЯ КЛИЕНТОВ');
    const pkData  = pkSheet.getDataRange().getValues();
    data.fieldIds.forEach(fid => {
      for (let i = 1; i < pkData.length; i++) {
        if (String(pkData[i][0]) === String(fid) && pkData[i][9] !== 'готово') {
          pkSheet.getRange(i + 1, 10).setValue('в работе');
        }
      }
    });
  }

  sendTelegram(GROUP_CHAT_ID,
    `☀️ *День ${data.dayNum} открыт*\n\n` +
    `📍 ${data.direction || data.tripId}\n📅 ${now}\n` +
    `🌾 Полей: ${(data.fieldIds||[]).length} · План: ${data.haPlan || 0} га`
  );

  return { ok: true, dayId: id };
}

function closeDay(data) {
  const sheet = getSheet('ДНИ');
  const found = findRowById(sheet, data.dayId);

  if (found) {
    sheet.getRange(found.rowNum, 6).setValue(data.haDone        || 0);
    sheet.getRange(found.rowNum, 7).setValue(data.fieldsDone    || 0);
    sheet.getRange(found.rowNum, 8).setValue(data.fieldsCarried || 0);
    sheet.getRange(found.rowNum, 9).setValue(data.spentToday    || 0);
    sheet.getRange(found.rowNum, 10).setValue('закрыт');
    if (data.note) sheet.getRange(found.rowNum, 11).setValue(data.note);
  }

  if (data.tripId) {
    const tripSheet = getSheet('КОМАНДИРОВКИ');
    const tripFound = findRowById(tripSheet, data.tripId);
    if (tripFound) {
      const adv = parseFloat(tripFound.rowData[6]) || 0;
      tripSheet.getRange(tripFound.rowNum, 8).setValue(data.totalSpent || 0);
      tripSheet.getRange(tripFound.rowNum, 9).setValue(adv - (data.totalSpent || 0));
      tripSheet.getRange(tripFound.rowNum, 10).setValue(data.totalHa || 0);
    }
  }

  const carried = data.fieldsCarried > 0 ? `\n⏩ Перенесено: ${data.fieldsCarried} пол.` : '';
  sendTelegram(GROUP_CHAT_ID,
    `🌙 *День ${data.dayNum} закрыт*\n\n📅 ${data.date}\n` +
    `🌾 Обработано: *${data.haDone} га*\n` +
    `✅ Полей закрыто: ${data.fieldsDone}${carried}\n` +
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
    id, data.tripId||'', data.dayId||'', formatDate(now), formatTime(now),
    data.label||data.cat||'', parseFloat(normalizeNumber(data.amount))||0, data.comment||''
  ]);

  if (data.tripId && data.totalSpent !== undefined) {
    const tripSheet = getSheet('КОМАНДИРОВКИ');
    const found     = findRowById(tripSheet, data.tripId);
    if (found) {
      const adv = parseFloat(found.rowData[6]) || 0;
      const sp  = parseFloat(data.totalSpent)  || 0;
      tripSheet.getRange(found.rowNum, 8).setValue(sp);
      tripSheet.getRange(found.rowNum, 9).setValue(adv - sp);
    }
  }

  return { id, ok: true };
}

function getExpenses(data) {
  const sheet = getSheet('РАСХОДЫ');
  const rows  = sheetRows(sheet);
  const expenses = rows.map(r => ({
    id: String(r[0]), tripId: String(r[1]), dayId: String(r[2]),
    date: r[3], time: r[4], label: r[5], amount: r[6], comment: r[7]
  })).filter(e => !data.tripId || String(e.tripId) === String(data.tripId));
  return { expenses, total: expenses.length };
}

// ══════════════════════════════════════════════════════════
// ОБРАБОТКИ
// ══════════════════════════════════════════════════════════

function addObrabotka(data) {
  const sheet = getSheet('ОБРАБОТКИ');
  const now   = formatDate(new Date());
  const id    = 'OB-' + Date.now();
  const ha    = parseFloat(normalizeNumber(data.haFact)) || 0;

  const opsArr  = data.operators ? data.operators.split(',').map(s => s.trim()).filter(Boolean) : [];
  const haPerOp = opsArr.length > 0 ? Math.round(ha / opsArr.length) : ha;

  sheet.appendRow([
    id, now, data.zayavkaId||'', data.fieldName||'', data.crop||'', data.prep||'',
    normalizeNumber(data.norm)||'', ha, data.bort||'', data.operators||'',
    haPerOp, data.photoUrl||'', 'закрыто', now, data.note||''
  ]);

  if (opsArr.length > 0) updateOperatorStats({ ids: opsArr, haPerOp });

  return { id, ok: true };
}

function getObrabotki(data) {
  const sheet = getSheet('ОБРАБОТКИ');
  const rows  = sheetRows(sheet);
  const list  = rows.map(r => ({
    id: String(r[0]), date: r[1], zayavkaId: r[2], fieldName: r[3],
    crop: r[4], prep: r[5], norm: r[6], haFact: r[7], bort: r[8],
    operators: r[9], haPerOp: r[10], photoUrl: r[11], status: r[12],
    dateClosed: r[13], note: r[14]
  }));
  return { obrabotki: list, total: list.length };
}

// ══════════════════════════════════════════════════════════
// ОПЕРАТОРЫ
// ══════════════════════════════════════════════════════════

function getOperators(data) {
  const sheet = getSheet('ОПЕРАТОРЫ');
  const rows  = sheetRows(sheet);
  const ops   = rows.map(r => ({
    id: String(r[0]), name: r[1], phone: r[2], region: r[3],
    cert: r[4], hours: r[5], status: r[6], regDate: r[7],
    telegramId: String(r[8]), type: r[9], hectares: r[10],
    trips: r[11], role: r[12], lastTrip: r[13]
  })).filter(o => !data.status || o.status === data.status);
  return { operators: ops, total: ops.length };
}

function updateOperatorStats(data) {
  const sheet = getSheet('ОПЕРАТОРЫ');
  const rows  = sheet.getDataRange().getValues();
  const now   = formatDate(new Date());

  (data.ids || []).forEach(opId => {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(opId)) {
        const curHa    = parseFloat(rows[i][10]) || 0;
        const curTrips = parseFloat(rows[i][11]) || 0;
        sheet.getRange(i+1, 11).setValue(curHa + (data.haPerOp || 0));
        sheet.getRange(i+1, 12).setValue(curTrips + 1);
        sheet.getRange(i+1, 14).setValue(now);
        break;
      }
    }
  });
  return { ok: true };
}

// ══════════════════════════════════════════════════════════
// ЗАЯВКИ
// ══════════════════════════════════════════════════════════

function getZayavki(data) {
  const sheet = getSheet('ЗАЯВКИ');
  const rows  = sheetRows(sheet);
  const list  = rows.map(r => ({
    id: String(r[0]), date: r[1], org: r[2], unp: r[3],
    phone: r[4], telegramId: String(r[5]), region: r[6],
    area: r[7], crop: r[8], coefCrop: r[9], complexity: r[10],
    obstacles: r[11], sections: r[12], urgency: r[13], dates: r[14],
    comment: r[15], status: r[16], dateStatus: r[17],
    coefFinal: r[18], priceWork: r[19], priceRoad: r[20],
    priceTotal: r[21], pricePerHa: r[22], managerNote: r[23], distance: r[24]
  })).filter(z => !data.status || z.status === data.status);
  return { zayavki: list, total: list.length };
}

function updateZayavka(data) {
  const sheet = getSheet('ЗАЯВКИ');
  const found = findRowById(sheet, data.id);
  if (!found) return { error: 'Заявка не найдена: ' + data.id };
  const now = formatDate(new Date());
  if (data.status)      sheet.getRange(found.rowNum, 17).setValue(data.status);
  if (data.status)      sheet.getRange(found.rowNum, 18).setValue(now);
  if (data.managerNote) sheet.getRange(found.rowNum, 24).setValue(data.managerNote);
  return { ok: true };
}

// ══════════════════════════════════════════════════════════
// КАНДИДАТЫ
// ══════════════════════════════════════════════════════════

function getKandidaty(data) {
  const sheet = getSheet('КАНДИДАТЫ');
  const rows  = sheetRows(sheet);
  const list  = rows.map(r => ({
    id: String(r[0]), date: r[1], name: r[2], age: r[3],
    region: r[4], phone: r[5], role: r[6], exp: r[7],
    car: r[8], trips: r[9], source: r[10], about: r[11],
    telegramId: String(r[12]), status: r[13], dateStatus: r[14], note: r[15]
  })).filter(k => !data.status || k.status === data.status);
  return { kandidaty: list, total: list.length };
}

function updateKandidat(data) {
  const sheet = getSheet('КАНДИДАТЫ');
  const found = findRowById(sheet, data.id);
  if (!found) return { error: 'Кандидат не найден: ' + data.id };
  const now = formatDate(new Date());
  if (data.status) sheet.getRange(found.rowNum, 14).setValue(data.status);
  if (data.status) sheet.getRange(found.rowNum, 15).setValue(now);
  if (data.note)   sheet.getRange(found.rowNum, 16).setValue(data.note);
  return { ok: true };
}

// ══════════════════════════════════════════════════════════
// БРИГАДЫ
// ══════════════════════════════════════════════════════════

function getBrigady(data) {
  const sheet = getSheet('БРИГАДЫ');
  const rows  = sheetRows(sheet);
  const list  = rows.map(r => ({
    id: String(r[0]), dateCreated: r[1], leaderId: String(r[2]),
    operatorIds: r[3], technicId: String(r[4]), tripId: String(r[5]),
    status: r[6], dateEnd: r[7], note: r[8]
  })).filter(b => !data.status || b.status === data.status);
  return { brigady: list, total: list.length };
}

// ══════════════════════════════════════════════════════════
// TELEGRAM
// ══════════════════════════════════════════════════════════

function sendTelegram(chatId, text) {
  if (!chatId || String(chatId).startsWith('ВАШ')) return;
  try {
    UrlFetchApp.fetch(
      'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage',
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ chat_id: chatId, text: text, parse_mode: 'Markdown' })
      }
    );
  } catch(e) {
    Logger.log('Telegram error: ' + e.message);
  }
}
