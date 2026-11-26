const ExcelJS = require("exceljs");

const SERVICE_URL = "https://roszdravnadzor.gov.ru/services/turnover";
const AJAX_URL = "https://roszdravnadzor.gov.ru/ajax/services/turnover";
const PAGE_SIZE = 500;
const MAX_RETRIES = 4;
const RETRY_DELAYS_MS = [1000, 2000, 4000];
const FETCH_TIMEOUT_MS = 20000;

const COLUMN_NAMES = [
  "Поисковый запрос",
  "Тип ЛС",
  "Дата внесения в АИС Росздравнадзора",
  "№ РУ",
  "Дата РУ",
  "Торговое наименование",
  "Производитель (выпускающий контроль)",
  "Страна",
  "Сведения о стадиях производства",
  "Нормативная документация",
  "Организация, выпустившая в гражданский оборот",
  "№ серии, партии",
  "№ разрешения, дата разрешения Росздравнадзора на ИЛП",
  "Письма, приказы об изъятии ЛС (номер, дата)"
];

let diagnosticLog = [];

function resetDiagnosticLog() {
  diagnosticLog = [];
}

function log(message, data = undefined) {
  const stamp = new Date().toISOString();
  let line = `[${stamp}] ${message}`;

  if (data !== undefined) {
    let serialized;
    try {
      serialized = JSON.stringify(data, null, 2);
    } catch {
      serialized = String(data);
    }
    line += `\n${serialized}`;
  }

  diagnosticLog.push(line);
  console.log(line);
}

function getDiagnosticLog() {
  return diagnosticLog.join("\n");
}

function truncate(value, maxLength = 3000) {
  value = String(value ?? "");
  return value.length <= maxLength
    ? value
    : value.slice(0, maxLength) + "... [TRUNCATED]";
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function serializeError(error) {
  const result = {
    name: error?.name ?? null,
    message: error?.message ?? String(error),
    code: error?.code ?? null,
    stack: error?.stack ?? null,
    cause: null
  };

  if (error?.cause) {
    const cause = error.cause;
    result.cause = {
      name: cause?.name ?? null,
      message: cause?.message ?? String(cause),
      code: cause?.code ?? null,
      errno: cause?.errno ?? null,
      syscall: cause?.syscall ?? null,
      address: cause?.address ?? null,
      port: cause?.port ?? null,
      hostname: cause?.hostname ?? null,
      stack: cause?.stack ?? null
    };
  }

  return result;
}

function retryableStatus(status) {
  return status === 408 || status === 429 || status >= 500;
}

async function fetchWithRetry(url, options, context) {
  let lastError = null;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);

    log(`${context}: attempt ${attempt}/${MAX_RETRIES}`, {
      url,
      method: options.method || "GET",
      timeout_ms: FETCH_TIMEOUT_MS
    });

    try {
      const response = await fetch(url, {
        ...options,
        signal: controller.signal
      });

      clearTimeout(timeout);

      log(`${context}: HTTP response`, {
        attempt,
        status: response.status,
        status_text: response.statusText
      });

      if (!retryableStatus(response.status) || attempt === MAX_RETRIES) {
        return response;
      }

      // Consume body before retry so the connection is released cleanly.
      const retryBody = await response.text();
      log(`${context}: retryable HTTP ${response.status}`, {
        attempt,
        response_body: truncate(retryBody, 1200)
      });
    } catch (error) {
      clearTimeout(timeout);
      lastError = error;

      log(`${context}: fetch exception`, {
        attempt,
        error: serializeError(error)
      });

      if (attempt === MAX_RETRIES) {
        throw error;
      }
    }

    const delay = RETRY_DELAYS_MS[Math.min(attempt - 1, RETRY_DELAYS_MS.length - 1)];
    log(`${context}: retry scheduled`, {
      delay_ms: delay,
      next_attempt: attempt + 1
    });
    await sleep(delay);
  }

  throw lastError || new Error(`${context}: retry loop exhausted`);
}

function cleanHtml(value) {
  if (value == null) return "";

  return String(value)
    .replace(/<\s*br\s*\/?\s*>/gi, "\n")
    .replace(/<\/\s*(div|p|li)\s*>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;|&#160;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#039;|&#39;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/[ \t]+/g, " ")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function extractCell(record, columnNumber) {
  const cell = record[`col${columnNumber}`] ?? "";

  if (!cell || typeof cell !== "object") {
    return cleanHtml(cell);
  }

  const title = cleanHtml(cell.title ?? "");
  const label = cleanHtml(cell.label ?? "");

  return title || label;
}

function buildPayload({
  tradeName,
  dateFrom,
  dateTo,
  start,
  draw,
  previousTotal
}) {
  const payload = new URLSearchParams({
    draw: String(draw),
    start: String(start),
    length: String(PAGE_SIZE),
    "search[value]": "",
    "search[regex]": "false",
    prev_total: String(previousTotal),
    q_label: "",
    q_series: "",
    q_type_ls: "2",
    q_mnn: "",
    q_dt_ru_from: "",
    q_dt_ru_to: "",
    q_no_ru: "",
    q_country: "",
    q_tn: tradeName,
    dt_from: dateFrom,
    dt_to: dateTo,
    q_producer: "",
    q_inf_letter: "0",
    q_org: ""
  });

  for (let i = 0; i < 13; i++) {
    payload.set(`columns[${i}][data]`, `col${i + 1}.label`);
    payload.set(`columns[${i}][name]`, "");
    payload.set(`columns[${i}][searchable]`, "true");
    payload.set(`columns[${i}][orderable]`, "true");
    payload.set(`columns[${i}][search][value]`, "");
    payload.set(`columns[${i}][search][regex]`, "false");
  }

  return payload;
}

function extractCookies(response) {
  let setCookies = [];

  if (typeof response.headers.getSetCookie === "function") {
    setCookies = response.headers.getSetCookie();
  } else {
    const raw = response.headers.get("set-cookie");
    if (raw) setCookies = [raw];
  }

  const pairs = setCookies
    .map(cookie => cookie.split(";")[0].trim())
    .filter(Boolean);

  return {
    count: setCookies.length,
    names: pairs.map(cookie => cookie.split("=")[0]),
    header: pairs.join("; ")
  };
}

const BROWSER_HEADERS = {
  "User-Agent":
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
    "AppleWebKit/537.36 (KHTML, like Gecko) " +
    "Chrome/131.0.0.0 Safari/537.36",
  "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
  "Cache-Control": "no-cache",
  "Pragma": "no-cache"
};

async function openService() {
  log("Opening Roszdravnadzor service", { url: SERVICE_URL });

  const started = Date.now();

  let response;
  try {
    response = await fetchWithRetry(
      SERVICE_URL,
      {
        method: "GET",
        redirect: "follow",
        headers: {
          ...BROWSER_HEADERS,
          "Accept":
            "text/html,application/xhtml+xml," +
            "application/xml;q=0.9,*/*;q=0.8"
        }
      },
      "Initial GET"
    );
  } catch (error) {
    log("Initial GET exhausted retries", {
      elapsed_ms: Date.now() - started,
      error: serializeError(error)
    });
    throw new Error(
      `Не удалось подключиться к Росздравнадзору после ${MAX_RETRIES} попыток: ` +
      `${error?.message || String(error)}`
    );
  }

  const text = await response.text();

  log("Initial GET completed", {
    status: response.status,
    elapsed_ms: Date.now() - started,
    final_url: response.url,
    content_type: response.headers.get("content-type"),
    bytes: text.length
  });

  if (!response.ok) {
    log("Initial GET failed", {
      status: response.status,
      response_body: truncate(text)
    });
    throw new Error(
      `Первичный GET Росздравнадзора: HTTP ${response.status}`
    );
  }

  const cookies = extractCookies(response);

  log("Cookies received", {
    count: cookies.count,
    names: cookies.names
  });

  log("Initial HTML received", {
    has_turnover_form: text.includes("id_q_tn")
  });

  return cookies.header;
}

async function requestPage({
  tradeName,
  dateFrom,
  dateTo,
  start,
  draw,
  previousTotal,
  cookieHeader
}) {
  const requestArgs = {
    tradeName,
    dateFrom,
    dateTo,
    start,
    draw,
    previousTotal
  };

  const payload = buildPayload(requestArgs);

  const headers = {
    ...BROWSER_HEADERS,
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "Origin": "https://roszdravnadzor.gov.ru",
    "Referer": SERVICE_URL,
    "X-Requested-With": "XMLHttpRequest"
  };

  if (cookieHeader) {
    headers.Cookie = cookieHeader;
  }

  log("Roszdravnadzor AJAX request", {
    trade_name: tradeName,
    date_from: dateFrom,
    date_to: dateTo,
    start,
    length: PAGE_SIZE,
    draw,
    prev_total: previousTotal,
    cookie_attached: Boolean(cookieHeader)
  });

  let response;
  try {
    response = await fetchWithRetry(
      AJAX_URL,
      {
        method: "POST",
        headers,
        body: payload.toString(),
        redirect: "follow"
      },
      `AJAX ${tradeName} start=${start}`
    );
  } catch (error) {
    log("AJAX exhausted retries", {
      trade_name: tradeName,
      start,
      draw,
      error: serializeError(error)
    });
    throw new Error(
      `${tradeName}: сетевой сбой Росздравнадзора после ${MAX_RETRIES} попыток: ` +
      `${error?.message || String(error)}`
    );
  }

  const responseText = await response.text();

  log("Roszdravnadzor AJAX final response", {
    status: response.status,
    content_type: response.headers.get("content-type"),
    content_length: responseText.length,
    start,
    draw
  });

  if (!response.ok) {
    log("Roszdravnadzor AJAX HTTP ERROR", {
      status: response.status,
      status_text: response.statusText,
      trade_name: tradeName,
      date_from: dateFrom,
      date_to: dateTo,
      start,
      draw,
      prev_total: previousTotal,
      response_body: truncate(responseText)
    });

    throw new Error(
      `${tradeName}: Росздравнадзор AJAX HTTP ${response.status}`
    );
  }

  let json;
  try {
    json = JSON.parse(responseText);
  } catch {
    log("AJAX response is not JSON", {
      trade_name: tradeName,
      body: truncate(responseText)
    });
    throw new Error(
      `${tradeName}: Росздравнадзор вернул AJAX-ответ не в формате JSON`
    );
  }

  if (json.message) {
    throw new Error(
      `${tradeName}: Росздравнадзор: ${json.message}`
    );
  }

  return json;
}

async function loadTradeName(
  tradeName,
  dateFrom,
  dateTo,
  cookieHeader
) {
  log("Starting trade name", {
    trade_name: tradeName
  });

  let all = [];
  let start = 0;
  let draw = 1;
  let total = null;
  let previousTotal = 0;

  while (true) {
    const json = await requestPage({
      tradeName,
      dateFrom,
      dateTo,
      start,
      draw,
      previousTotal,
      cookieHeader
    });

    if (total === null) {
      total = Number(
        json.recordsFiltered ??
        json.recordsTotal ??
        0
      );

      previousTotal = Number(
        json.recordsTotal ??
        total
      );

      log("Search initialized", {
        trade_name: tradeName,
        records_filtered: json.recordsFiltered,
        records_total: json.recordsTotal,
        expected: total
      });
    }

    const batch = json.data || [];

    log("Batch received", {
      trade_name: tradeName,
      start,
      batch: batch.length,
      accumulated: all.length + batch.length,
      expected: total
    });

    if (!batch.length) {
      break;
    }

    all.push(...batch);
    start += batch.length;
    draw += 1;

    if (all.length >= total) {
      break;
    }
  }

  if (all.length !== total) {
    throw new Error(
      `${tradeName}: загружено ${all.length} из ${total}`
    );
  }

  log("Trade name completed", {
    trade_name: tradeName,
    records: all.length
  });

  return {
    tradeName,
    count: all.length,
    rows: all.map(record => [
      tradeName,
      ...Array.from(
        { length: 13 },
        (_, index) => extractCell(record, index + 1)
      )
    ])
  };
}

function safeFilename(names) {
  const base = names
    .join("_")
    .replace(/[\\/:*?"<>|]+/g, "_")
    .replace(/\s+/g, " ")
    .slice(0, 120) || "Результат";

  return `${base}_Росздравнадзор_ЛП.xlsx`;
}

async function createWorkbook(rows) {
  const workbook = new ExcelJS.Workbook();

  const worksheet = workbook.addWorksheet(
    "Результаты",
    {
      views: [
        {
          state: "frozen",
          ySplit: 1
        }
      ]
    }
  );

  worksheet.addRow(COLUMN_NAMES);

  rows.forEach(row => worksheet.addRow(row));

  worksheet.autoFilter = {
    from: "A1",
    to: `N${Math.max(1, rows.length + 1)}`
  };

  const header = worksheet.getRow(1);

  header.font = {
    bold: true
  };

  header.alignment = {
    horizontal: "center",
    vertical: "middle",
    wrapText: true
  };

  header.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: {
      argb: "FFD9EAF7"
    }
  };

  const widths = [
    22, 12, 20, 20, 13, 28, 38,
    18, 48, 40, 42, 22, 32, 34
  ];

  widths.forEach((width, index) => {
    worksheet.getColumn(index + 1).width = width;
  });

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      row.alignment = {
        vertical: "top",
        wrapText: true
      };
    }
  });

  return workbook.xlsx.writeBuffer();
}

exports.handler = async function handler(event) {
  resetDiagnosticLog();

  const invocationId = Math.random()
    .toString(36)
    .slice(2, 10);

  log("FUNCTION START", {
    invocation_id: invocationId,
    method: event.httpMethod,
    node: process.version,
    region: process.env.AWS_REGION || null,
    max_retries: MAX_RETRIES,
    retry_delays_ms: RETRY_DELAYS_MS,
    fetch_timeout_ms: FETCH_TIMEOUT_MS
  });

  if (event.httpMethod !== "POST") {
    return {
      statusCode: 405,
      headers: {
        "Content-Type": "application/json; charset=utf-8"
      },
      body: JSON.stringify({
        error: "Method not allowed"
      })
    };
  }

  try {
    const body = JSON.parse(event.body || "{}");

    const names = [
      ...new Set(
        (body.trade_names || [])
          .map(value => String(value).trim())
          .filter(Boolean)
      )
    ];

    if (!names.length) {
      throw new Error(
        "Укажите минимум одно торговое наименование"
      );
    }

    if (!body.date_from || !body.date_to) {
      throw new Error("Укажите обе даты");
    }

    log("Request accepted", {
      invocation_id: invocationId,
      trade_names: names,
      date_from: body.date_from,
      date_to: body.date_to
    });

    /*
      Одна сессия Росздравнадзора на весь пользовательский запрос.
    */
    log("Opening shared Roszdravnadzor session");

    const cookieHeader = await openService();

    log("Shared Roszdravnadzor session ready", {
      trade_names: names.length,
      cookie_attached: Boolean(cookieHeader)
    });

    const searchResults = [];
    let rows = [];

    for (const name of names) {
      const result = await loadTradeName(
        name,
        body.date_from,
        body.date_to,
        cookieHeader
      );

      searchResults.push({
        trade_name: result.tradeName,
        count: result.count
      });

      log("Trade name search result", {
        trade_name: result.tradeName,
        count: result.count
      });

      if (result.count > 0) {
        rows.push(...result.rows);
      }
    }

    log("Search summary", {
      search_results: searchResults,
      total_rows_before_deduplication: rows.length
    });

    /*
      Все ТН дали 0: XLSX не создаём.
    */
    if (rows.length === 0) {
      log("NO RESULTS FOR ALL TRADE NAMES", {
        search_results: searchResults
      });

      return {
        statusCode: 404,
        headers: {
          "Content-Type": "application/json; charset=utf-8",
          "Cache-Control": "no-store"
        },
        body: JSON.stringify({
          no_results: true,
          message: "По заданным параметрам сведения не найдены.",
          search_results: searchResults,
          invocation_id: invocationId
        })
      };
    }

    /*
      Полные дубликаты удаляем после объединения.
      Первая колонка содержит поисковое ТН, поэтому данные разных
      поисковых запросов не сольются случайно между собой.
    */
    const seen = new Set();

    rows = rows.filter(row => {
      const key = JSON.stringify(row);

      if (seen.has(key)) {
        return false;
      }

      seen.add(key);
      return true;
    });

    log("Creating workbook", {
      rows_after_deduplication: rows.length,
      search_results: searchResults
    });

    const buffer = await createWorkbook(rows);

    const foundNames = searchResults
      .filter(item => item.count > 0)
      .map(item => item.trade_name);

    const filename = safeFilename(foundNames);

    log("FUNCTION SUCCESS", {
      invocation_id: invocationId,
      filename,
      rows: rows.length,
      bytes: buffer.length,
      search_results: searchResults
    });

    return {
      statusCode: 200,
      isBase64Encoded: true,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition":
          `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
        "X-Filename": encodeURIComponent(filename),
        "X-Search-Results": encodeURIComponent(
          JSON.stringify(searchResults)
        ),
        "Cache-Control": "no-store"
      },
      body: Buffer.from(buffer).toString("base64")
    };
  } catch (error) {
    log("FUNCTION ERROR", {
      invocation_id: invocationId,
      error: serializeError(error)
    });

    return {
      statusCode: 500,
      headers: {
        "Content-Type": "application/json; charset=utf-8",
        "Cache-Control": "no-store"
      },
      body: JSON.stringify({
        error: error?.message || String(error),
        invocation_id: invocationId,
        diagnostic_log: getDiagnosticLog()
      })
    };
  }
};
