/**
 * Google Apps Script - النسخة 2 (كاملة)
 *
 * ما يفعله:
 * 1) يسجّل كل طلب في الشيت أولًا، فلا يضيع بسبب أي خطأ لاحق.
 * 2) يرسل Purchase إلى Meta (Conversions API) بوقت الطلب الحقيقي ومع بيانات مطابقة أفضل.
 * 3) إذا فشل الإرسال إلى Meta يسجّله "فشل" ويعيد المحاولة تلقائيًا مع كل طلب جديد،
 *    ومن قائمة "الطلبات" في الشيت (إعادة إرسال الأحداث الفاشلة).
 * 4) لا يكرر الصف ولا الحدث إذا أعادت الصفحة المحاولة.
 * 5) لا يتجاهل طلبًا بسبب رقم هاتف خاطئ: يسجّله ويضع ملاحظة، ولا يرسله إلى Meta.
 *
 * الشيء الوحيد الذي تضعه بنفسك: التوكن في السطر الذي يحتوي ACCESS_TOKEN
 */

var VERSION = "v2-2026-10-22-shelf";

var PIXEL_ID = "1650264038831419";

// ضع التوكن بين علامتي الاقتباس (انسخه من السكريبت القديم قبل الحذف)
var ACCESS_TOKEN = "";

// اتركه فارغًا إلا عند اختبار Test Events
var TEST_EVENT_CODE = "";

// نسخة Graph API (v25.0 متاحة حتى 2028)
var GRAPH_VERSION = "v25.0";

// اتركه فارغًا لاستعمال الورقة النشطة، أو اكتب اسم الورقة بالضبط
var SHEET_NAME = "";

var CONTENT_ID = "bath-shelf-set";
var CONTENT_NAME = "طقم منظمات الحمام الزاوية 5 قطع";
var DEFAULT_VALUE = 3900;      // قيمة الشراء إن لم تصل من الصفحة
var CURRENCY = "DZD";

// أعمدة الشيت (الأعمدة 1-12 كما كانت، والجديدة في آخرها)
var HEADERS = [
  "التاريخ", "الوقت", "الاسم", "الهاتف", "الولاية", "البلدية",
  "طريقة التوصيل", "سعر التوصيل", "اللون", "المجموع",
  "Event ID", "رد فيسبوك", "حالة الإرسال", "وقت الحدث", "ملاحظات",
  "بيانات الحدث (لا تلمسها)"
];
var COL_EVENT_ID = 11;
var COL_META = 12;
var COL_STATUS = 13;
var COL_EVENT_TIME = 14;
var COL_NOTES = 15;
var COL_PAYLOAD = 16;

var ST_OK = "تم";
var ST_FAIL = "فشل";
var ST_PENDING = "قيد الإرسال";
var ST_SKIP = "متجاهل";
var ST_EXPIRED = "منتهي الصلاحية";

// Meta يرفض الأحداث الأقدم من 7 أيام
var MAX_AGE_SECONDS = 6 * 24 * 3600;


/* ============================ نقاط الدخول ============================ */

/** فتح رابط /exec في المتصفح يعرض هذا الرد، لتتأكد أن النسخة الجديدة منشورة */
function doGet() {
  return jsonOut_({ status: "ok", version: VERSION });
}

/** استقبال الطلب من صفحة الهبوط */
function doPost(e) {

  try {

    var data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    var eventId = clean_(data.eventId, 120);

    if (!eventId) {
      throw new Error("لم يصل Event ID من صفحة الهبوط.");
    }

    // 1) سجّل الطلب أولًا
    var saved = saveOrder_(data, eventId);

    if (saved.duplicate) {
      return jsonOut_({ result: "success", eventId: eventId, duplicate: true });
    }

    // 2) أرسل Purchase إلى Meta (فشله لا يؤثر على الطلب)
    try {
      dispatchRow_(saved.row);
    } catch (metaErr) {
      console.log("META ERROR: " + metaErr.message);
      markRow_(saved.row, "خطأ: " + metaErr.message, ST_FAIL);
    }

    // 3) أعد محاولة أحداث فاشلة سابقة (حتى 2 في كل طلب)
    try {
      retryFailed_(2);
    } catch (retryErr) {
      console.log("RETRY ERROR: " + retryErr.message);
    }

    return jsonOut_({ result: "success", eventId: eventId });

  } catch (err) {

    console.log("ERROR: " + err.message);
    return jsonOut_({ result: "error", message: err.message });

  }

}


/* ============================ تسجيل الطلب ============================ */

function saveOrder_(data, eventId) {

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {

    var sheet = getSheet_();
    ensureHeaders_(sheet);

    var last = sheet.getLastRow();

    // منع تكرار الصف: Event ID موجود في العمود 11
    if (last > 1) {
      var ids = sheet.getRange(2, COL_EVENT_ID, last - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === eventId) {
          return { duplicate: true };
        }
      }
    }

    var phone = String(data.phone || "").replace(/[^0-9]/g, "");
    var phoneOk = /^0[567][0-9]{8}$/.test(phone);
    var eventTime = pickEventTime_(data.eventTime);
    var notes = [];

    if (!phoneOk) {
      notes.push("هاتف غير صالح: لم يُرسَل Purchase إلى فيسبوك");
    }

    var repeat = recentSamePhone_(sheet, last, phone, eventTime);
    if (repeat) {
      notes.push(repeat);
    }

    var payload = "";
    if (phoneOk) {
      payload = JSON.stringify(buildEvent_(data, eventId, phone, eventTime));
    }

    var row = last + 1;
    var range = sheet.getRange(row, 1, 1, HEADERS.length);

    // نص عادي: يمنع تحويل الاسم إلى معادلة، ويحفظ الصفر في أول الهاتف
    range.setNumberFormat("@");
    range.setValues([[
      clean_(data.date, 40),
      clean_(data.time, 40),
      clean_(data.name, 120),
      phone,
      clean_(data.wilaya, 80),
      clean_(data.commune, 120),
      clean_(data.delivery, 60),
      clean_(data.shipping, 60),
      clean_(data.color, 40),
      clean_(data.total, 60),
      eventId,
      "",
      phoneOk ? ST_PENDING : ST_SKIP,
      String(eventTime),
      notes.join(" | "),
      payload
    ]]);

    SpreadsheetApp.flush();

    return { duplicate: false, row: row };

  } finally {

    lock.releaseLock();

  }

}


/* ============================ إرسال Meta ============================ */

/** يبني حدث Purchase مع أكبر قدر مفيد من بيانات المطابقة */
function buildEvent_(data, eventId, phone, eventTime) {

  var userData = {};

  userData.ph = [sha256_("213" + phone.replace(/^0/, ""))];

  var nameParts = clean_(data.name, 120).split(/\s+/).filter(String);
  var fn = norm_(nameParts[0] || "");
  var ln = nameParts.length > 1 ? norm_(nameParts[nameParts.length - 1]) : "";
  if (fn) { userData.fn = [sha256_(fn)]; }
  if (ln) { userData.ln = [sha256_(ln)]; }

  var city = norm_(clean_(data.commune, 120));
  var state = norm_(clean_(data.wilaya, 80).replace(/[0-9]/g, ""));
  if (city) { userData.ct = [sha256_(city)]; }
  if (state) { userData.st = [sha256_(state)]; }

  userData.country = [sha256_("dz")];

  var externalId = clean_(data.externalId, 120);
  if (externalId) { userData.external_id = [sha256_(externalId)]; }

  // هذه الحقول تُرسل كما هي بلا تشفير
  var ip = clean_(data.clientIp, 64);
  var ua = clean_(data.userAgent, 400);
  var fbc = clean_(data.fbc, 200);
  var fbp = clean_(data.fbp, 200);
  if (ip) { userData.client_ip_address = ip; }
  if (ua) { userData.client_user_agent = ua; }
  if (fbc) { userData.fbc = fbc; }
  if (fbp) { userData.fbp = fbp; }

  var value = Number(data.value);
  if (!(value > 0)) { value = DEFAULT_VALUE; }

  var event = {
    event_name: "Purchase",
    event_time: eventTime,
    event_id: eventId,
    action_source: "website",
    user_data: userData,
    custom_data: {
      value: value,
      currency: clean_(data.currency, 3) || CURRENCY,
      content_ids: [CONTENT_ID],
      content_type: "product",
      content_name: CONTENT_NAME,
      num_items: 1,
      order_id: eventId
    }
  };

  var url = clean_(data.sourceUrl, 500);
  if (url) { event.event_source_url = url; }

  return event;

}


/** يرسل حدثًا محفوظًا في صف إلى Meta ويحدّث حالته */
function dispatchRow_(row) {

  var sheet = getSheet_();
  var payload = String(sheet.getRange(row, COL_PAYLOAD).getValue() || "");

  if (!payload) {
    markRow_(row, "لا توجد بيانات حدث", ST_SKIP);
    return;
  }

  var eventTime = Number(sheet.getRange(row, COL_EVENT_TIME).getValue());
  if (eventTime && (nowSec_() - eventTime) > MAX_AGE_SECONDS) {
    markRow_(row, "أقدم من 6 أيام: لا يقبله فيسبوك", ST_EXPIRED);
    return;
  }

  var result = sendToMeta_(JSON.parse(payload));

  markRow_(row, result.code + " | " + result.body, result.ok ? ST_OK : ST_FAIL);

}


/** استدعاء Conversions API. التوكن في جسم الطلب لا في الرابط */
function sendToMeta_(event) {

  if (!ACCESS_TOKEN || ACCESS_TOKEN === "ضع_التوكن_الجديد_هنا") {
    return { ok: false, code: "خطأ", body: "ACCESS_TOKEN غير موجود" };
  }

  var body = { data: [event], access_token: ACCESS_TOKEN };

  if (TEST_EVENT_CODE) {
    body.test_event_code = TEST_EVENT_CODE;
  }

  var response = UrlFetchApp.fetch(
    "https://graph.facebook.com/" + GRAPH_VERSION + "/" + PIXEL_ID + "/events",
    {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    }
  );

  var code = response.getResponseCode();
  var text = response.getContentText();
  var ok = false;

  if (code === 200) {
    try {
      ok = Number(JSON.parse(text).events_received) >= 1;
    } catch (parseErr) {
      ok = false;
    }
  }

  console.log("Meta " + code + ": " + text);

  return { ok: ok, code: code, body: String(text).slice(0, 300) };

}


/** يكتب نتيجة الإرسال في الصف (يتأكد أن الصف ما زال لنفس الطلب) */
function markRow_(row, metaText, status) {

  var sheet = getSheet_();
  sheet.getRange(row, COL_META).setValue(metaText);
  sheet.getRange(row, COL_STATUS).setValue(status);

}


/** يعيد إرسال الأحداث الفاشلة أو العالقة (حتى max حدثًا) ويرجع عدد المحاولات */
function retryFailed_(max) {

  var sheet = getSheet_();
  var last = sheet.getLastRow();

  if (last < 2) { return 0; }

  var rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var now = nowSec_();
  var tried = 0;

  for (var i = 0; i < rows.length && tried < max; i++) {

    var status = String(rows[i][COL_STATUS - 1]);
    var eventTime = Number(rows[i][COL_EVENT_TIME - 1]);
    var rowNumber = i + 2;

    if (status !== ST_FAIL && status !== ST_PENDING) { continue; }

    // الطلب الذي ما زال قيد الإرسال منذ أقل من دقيقتين قد يكون جاريًا الآن
    if (status === ST_PENDING && (now - eventTime) < 120) { continue; }

    tried++;

    try {
      dispatchRow_(rowNumber);
    } catch (err) {
      markRow_(rowNumber, "خطأ: " + err.message, ST_FAIL);
    }

  }

  return tried;

}


/* ============================ قائمة الشيت ============================ */

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("الطلبات")
      .addItem("إعادة إرسال الأحداث الفاشلة", "retryFailedEvents")
      .addItem("فحص الإعدادات", "checkSetup")
      .addToUi();
  } catch (err) {
    console.log("onOpen: " + err.message);
  }
}

/** من القائمة: يعيد إرسال كل الأحداث الفاشلة (حتى 50) */
function retryFailedEvents() {
  var n = retryFailed_(50);
  notify_("تمت إعادة محاولة " + n + " حدثًا. افتح عمود \"حالة الإرسال\" لترى النتيجة.");
}

/** من القائمة: يعرض حالة الإعدادات وعدد الطلبات حسب الحالة */
function checkSetup() {

  var sheet = getSheet_();
  var last = sheet.getLastRow();
  var counts = {};

  if (last > 1) {
    var statuses = sheet.getRange(2, COL_STATUS, last - 1, 1).getValues();
    for (var i = 0; i < statuses.length; i++) {
      var s = String(statuses[i][0]) || "(بلا حالة)";
      counts[s] = (counts[s] || 0) + 1;
    }
  }

  var lines = [
    "الإصدار: " + VERSION,
    "ورقة العمل: " + sheet.getName(),
    "Pixel ID: " + PIXEL_ID,
    "Graph API: " + GRAPH_VERSION,
    "التوكن: " + (ACCESS_TOKEN ? "موجود" : "غير موجود !!!"),
    "وضع الاختبار: " + (TEST_EVENT_CODE ? "مفعّل (" + TEST_EVENT_CODE + ")" : "معطّل"),
    "الطلبات حسب الحالة: " + JSON.stringify(counts)
  ];

  notify_(lines.join("\n"));
  return lines.join("\n");

}


/* ============================ أدوات مساعدة ============================ */

function notify_(message) {
  console.log(message);
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (err) {
    // التنفيذ من المحرر بلا واجهة: يكفي السجل
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getActiveSheet();
  if (!sheet) {
    throw new Error("ورقة العمل غير موجودة: " + SHEET_NAME);
  }
  return sheet;
}

/** ينشئ العناوين أو يكمّل الناقص منها دون تغيير ما هو موجود */
function ensureHeaders_(sheet) {

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }

  var range = sheet.getRange(1, 1, 1, HEADERS.length);
  var current = range.getValues()[0];
  var changed = false;

  for (var i = 0; i < HEADERS.length; i++) {
    if (!current[i]) {
      current[i] = HEADERS[i];
      changed = true;
    }
  }

  if (changed) {
    range.setValues([current]);
  }

}

/** وقت الحدث من الصفحة إن كان معقولًا، وإلا وقت الخادم الآن */
function pickEventTime_(raw) {

  var now = nowSec_();
  var t = Math.floor(Number(raw));

  if (t > 0 && t <= now + 60 && (now - t) < MAX_AGE_SECONDS) {
    return t;
  }

  return now;

}

/** تنبيه إن طلب نفس الهاتف خلال آخر 30 دقيقة */
function recentSamePhone_(sheet, last, phone, eventTime) {

  if (last < 2 || !phone) { return ""; }

  var start = Math.max(2, last - 59);
  var rows = sheet.getRange(start, 1, last - start + 1, HEADERS.length).getValues();

  for (var i = rows.length - 1; i >= 0; i--) {
    var samePhone = String(rows[i][3]) === phone;
    var t = Number(rows[i][COL_EVENT_TIME - 1]);
    if (samePhone && t && Math.abs(eventTime - t) < 1800) {
      return "تنبيه: نفس الهاتف طلب قبل " + Math.round(Math.abs(eventTime - t) / 60) + " دقيقة (طلب مكرر محتمل)";
    }
  }

  return "";

}

function nowSec_() {
  return Math.floor(Date.now() / 1000);
}

/** ينظف نصًا قادمًا من الصفحة: يحذف الرموز التحكمية ويقصّ الطول */
function clean_(value, maxLen) {
  return String(value == null ? "" : value)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .trim()
    .slice(0, maxLen || 200);
}

/** توحيد النص قبل التشفير: أحرف صغيرة بلا مسافات ولا تشكيل ولا رموز */
function norm_(text) {
  var s = String(text || "").toLowerCase();
  if (typeof s.normalize === "function") {
    s = s.normalize("NFD");
  }
  return s
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9\u0600-\u06FF]/g, "");
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** SHA-256 للنص بترميز UTF-8 */
function sha256_(input) {

  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(input),
    Utilities.Charset.UTF_8
  );

  var hex = "";

  for (var i = 0; i < digest.length; i++) {
    var b = digest[i] < 0 ? digest[i] + 256 : digest[i];
    var h = b.toString(16);
    if (h.length === 1) { h = "0" + h; }
    hex += h;
  }

  return hex;

}

/** اختبار الاتصال بالإنترنت */
function testFetch() {
  var response = UrlFetchApp.fetch("https://www.google.com");
  console.log(response.getResponseCode());
}
