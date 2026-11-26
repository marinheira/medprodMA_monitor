const ExcelJS = require("exceljs");

const SERVICE_URL =
  "https://roszdravnadzor.gov.ru/services/turnover";

const AJAX_URL =
  "https://roszdravnadzor.gov.ru/ajax/services/turnover";

const PAGE_SIZE = 500;

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

  const stamp =
    new Date().toISOString();


  let line =
    `[${stamp}] ${message}`;


  if (data !== undefined) {

    let serialized;

    try {
      serialized =
        JSON.stringify(
          data,
          null,
          2
        );
    } catch {
      serialized =
        String(data);
    }

    line +=
      "\n" +
      serialized;
  }


  diagnosticLog.push(
    line
  );


  console.log(
    line
  );
}


function getDiagnosticLog() {
  return diagnosticLog.join(
    "\n"
  );
}

function truncate(value, maxLength = 1500) {
  value = String(value ?? "");

  if (value.length <= maxLength) {
    return value;
  }

  return value.slice(0, maxLength) + "... [TRUNCATED]";
}


function cleanHtml(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value)
    .replace(
      /<\s*br\s*\/?\s*>/gi,
      "\n"
    )
    .replace(
      /<\/\s*(div|p|li)\s*>/gi,
      "\n"
    )
    .replace(
      /<[^>]+>/g,
      ""
    )
    .replace(
      /&nbsp;|&#160;/g,
      " "
    )
    .replace(
      /&amp;/g,
      "&"
    )
    .replace(
      /&quot;/g,
      "\""
    )
    .replace(
      /&#039;|&#39;/g,
      "'"
    )
    .replace(
      /&lt;/g,
      "<"
    )
    .replace(
      /&gt;/g,
      ">"
    )
    .replace(
      /[ \t]+/g,
      " "
    )
    .replace(
      /\n[ \t]+/g,
      "\n"
    )
    .replace(
      /\n{3,}/g,
      "\n\n"
    )
    .trim();
}


function extractCell(record, columnNumber) {
  const cell =
    record[`col${columnNumber}`] ?? "";

  if (
    !cell ||
    typeof cell !== "object"
  ) {
    return cleanHtml(cell);
  }

  const title =
    cleanHtml(cell.title ?? "");

  const label =
    cleanHtml(cell.label ?? "");

  // Подтверждено локальной диагностикой:
  // полный текст длинных значений находится в title.
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

  const payload =
    new URLSearchParams({
      draw: String(draw),
      start: String(start),
      length: String(PAGE_SIZE),

      "search[value]": "",
      "search[regex]": "false",

      prev_total:
        String(previousTotal),

      q_label: "",
      q_series: "",

      // 2 = ЛП
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


  for (
    let index = 0;
    index < 13;
    index++
  ) {

    const number =
      index + 1;

    payload.set(
      `columns[${index}][data]`,
      `col${number}.label`
    );

    payload.set(
      `columns[${index}][name]`,
      ""
    );

    payload.set(
      `columns[${index}][searchable]`,
      "true"
    );

    payload.set(
      `columns[${index}][orderable]`,
      "true"
    );

    payload.set(
      `columns[${index}][search][value]`,
      ""
    );

    payload.set(
      `columns[${index}][search][regex]`,
      "false"
    );
  }

  return payload;
}


function extractCookies(response) {

  /*
    В Node 20+/24 Headers может предоставлять getSetCookie().
    Это существенно лучше простого headers.get("set-cookie"),
    когда сервер устанавливает несколько cookies.
  */

  let setCookies = [];

  if (
    typeof response.headers.getSetCookie ===
    "function"
  ) {
    setCookies =
      response.headers.getSetCookie();
  } else {

    const raw =
      response.headers.get(
        "set-cookie"
      );

    if (raw) {
      setCookies = [raw];
    }
  }


  const cookiePairs =
    setCookies
      .map((cookie) => {

        // Из:
        // PHPSESSID=xxx; path=/; HttpOnly
        //
        // получаем:
        // PHPSESSID=xxx

        return cookie
          .split(";")[0]
          .trim();

      })
      .filter(Boolean);


  return {
    count: setCookies.length,

    names: cookiePairs.map(
      (cookie) =>
        cookie.split("=")[0]
    ),

    header:
      cookiePairs.join("; ")
  };
}


async function openService() {

  log(
    "Opening Roszdravnadzor service",
    {
      url: SERVICE_URL
    }
  );


  const started =
    Date.now();


  const response =
    await fetch(
      SERVICE_URL,
      {
        method: "GET",

        redirect: "follow",

        headers: {
          "User-Agent":
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
            "AppleWebKit/537.36 (KHTML, like Gecko) " +
            "Chrome/131.0.0.0 Safari/537.36",

          "Accept":
            "text/html,application/xhtml+xml," +
            "application/xml;q=0.9,*/*;q=0.8",

          "Accept-Language":
            "ru-RU,ru;q=0.9,en;q=0.8",

          "Cache-Control":
            "no-cache",

          "Pragma":
            "no-cache"
        }
      }
    );


  const elapsed =
    Date.now() - started;


  log(
    "Initial GET completed",
    {
      status: response.status,
      elapsed_ms: elapsed,
      final_url: response.url,
      content_type:
        response.headers.get(
          "content-type"
        )
    }
  );


  if (!response.ok) {

    const text =
      await response.text();

    log(
      "Initial GET failed",
      {
        status: response.status,
        body: truncate(text)
      }
    );


    throw new Error(
      `Первичный GET Росздравнадзора: HTTP ${response.status}`
    );
  }


  const cookies =
    extractCookies(response);


  log(
    "Cookies received",
    {
      count: cookies.count,
      names: cookies.names
    }
  );


  /*
    Тело читаем намеренно, чтобы запрос был
    полностью завершён до AJAX POST.
  */

  const html =
    await response.text();


  log(
    "Initial HTML received",
    {
      bytes: html.length,

      has_turnover_form:
        html.includes(
          "id_q_tn"
        )
    }
  );


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

  const payload =
    buildPayload({
      tradeName,
      dateFrom,
      dateTo,
      start,
      draw,
      previousTotal
    });


  log(
    "Roszdravnadzor AJAX request",
    {
      trade_name: tradeName,
      date_from: dateFrom,
      date_to: dateTo,
      start,
      length: PAGE_SIZE,
      draw,
      prev_total: previousTotal,

      cookie_attached:
        Boolean(cookieHeader)
    }
  );


  const headers = {

    "User-Agent":
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
      "AppleWebKit/537.36 (KHTML, like Gecko) " +
      "Chrome/131.0.0.0 Safari/537.36",

    "Accept":
      "application/json, " +
      "text/javascript, */*; q=0.01",

    "Accept-Language":
      "ru-RU,ru;q=0.9,en;q=0.8",

    "Content-Type":
      "application/x-www-form-urlencoded; charset=UTF-8",

    "Origin":
      "https://roszdravnadzor.gov.ru",

    "Referer":
      SERVICE_URL,

    "X-Requested-With":
      "XMLHttpRequest",

    "Cache-Control":
      "no-cache",

    "Pragma":
      "no-cache"
  };


  if (cookieHeader) {
    headers.Cookie =
      cookieHeader;
  }


  const started =
    Date.now();


  let response;

  try {

    response =
      await fetch(
        AJAX_URL,
        {
          method: "POST",
          headers,
          body:
            payload.toString(),

          redirect:
            "follow"
        }
      );

  } catch (error) {

    log(
      "AJAX network error",
      {
        message:
          error.message,

        name:
          error.name,

        cause:
          error.cause
            ? String(error.cause)
            : null
      }
    );

    throw error;
  }


  const elapsed =
    Date.now() - started;


  const responseText =
    await response.text();


  log(
    "Roszdravnadzor AJAX response",
    {
      status:
        response.status,

      elapsed_ms:
        elapsed,

      content_type:
        response.headers.get(
          "content-type"
        ),

      content_length:
        responseText.length,

      start,

      draw
    }
  );


  if (!response.ok) {

    /*
      Это самый важный диагностический лог.

      Он появится:
      Netlify
        -> Functions
        -> export
        -> Logs
    */

    log(
      "Roszdravnadzor AJAX HTTP ERROR",
      {
        status:
          response.status,

        status_text:
          response.statusText,

        trade_name:
          tradeName,

        date_from:
          dateFrom,

        date_to:
          dateTo,

        start,

        draw,

        prev_total:
          previousTotal,

        response_body:
          truncate(
            responseText,
            3000
          )
      }
    );


    throw new Error(
      `Росздравнадзор AJAX: HTTP ${response.status}`
    );
  }


  let json;

  try {

    json =
      JSON.parse(
        responseText
      );

  } catch {

    log(
      "AJAX response is not JSON",
      {
        body:
          truncate(
            responseText,
            3000
          )
      }
    );


    throw new Error(
      "Росздравнадзор вернул AJAX-ответ не в формате JSON"
    );
  }


  if (json.message) {

    log(
      "Roszdravnadzor message",
      {
        message:
          json.message
      }
    );


    throw new Error(
      `Росздравнадзор: ${json.message}`
    );
  }


  return json;
}


async function loadTradeName(
  tradeName,
  dateFrom,
  dateTo
) {

  log(
    "Starting trade name",
    {
      trade_name:
        tradeName
    }
  );


  /*
    Для КАЖДОГО торгового наименования создаём
    новую web-session Росздравнадзора.

    Это ближе всего к рабочему requests.Session()
    в Python.
  */

  const cookieHeader =
    await openService();


  let allRecords = [];

  let start = 0;
  let draw = 1;

  let total = null;
  let previousTotal = 0;


  while (true) {

    const json =
      await requestPage({
        tradeName,
        dateFrom,
        dateTo,
        start,
        draw,
        previousTotal,
        cookieHeader
      });


    if (total === null) {

      total =
        Number(
          json.recordsFiltered ??
          json.recordsTotal ??
          0
        );


      previousTotal =
        Number(
          json.recordsTotal ??
          total
        );


      log(
        "Search initialized",
        {
          trade_name:
            tradeName,

          records_filtered:
            json.recordsFiltered,

          records_total:
            json.recordsTotal,

          expected:
            total
        }
      );
    }


    const batch =
      json.data || [];


    log(
      "Batch received",
      {
        trade_name:
          tradeName,

        start,

        batch:
          batch.length,

        accumulated:
          allRecords.length +
          batch.length,

        expected:
          total
      }
    );


    if (!batch.length) {
      break;
    }


    allRecords.push(
      ...batch
    );


    start +=
      batch.length;

    draw += 1;


    if (
      allRecords.length >=
      total
    ) {
      break;
    }
  }


  if (
    allRecords.length !==
    total
  ) {

    throw new Error(
      `${tradeName}: загружено ` +
      `${allRecords.length} из ${total}`
    );
  }


  log(
    "Trade name completed",
    {
      trade_name:
        tradeName,

      records:
        allRecords.length
    }
  );


  return allRecords.map(
    (record) => [

      tradeName,

      ...Array.from(
        {
          length: 13
        },

        (_, index) =>
          extractCell(
            record,
            index + 1
          )
      )
    ]
  );
}


function safeFilename(names) {

  const base =
    names
      .join("_")
      .replace(
        /[\\/:*?"<>|]+/g,
        "_"
      )
      .replace(
        /\s+/g,
        " "
      )
      .slice(
        0,
        120
      ) ||
    "Результат";


  return (
    `${base}` +
    `_Росздравнадзор_ЛП.xlsx`
  );
}


async function createWorkbook(
  rows
) {

  const workbook =
    new ExcelJS.Workbook();


  workbook.creator =
    "Roszdravnadzor Turnover Export";


  const worksheet =
    workbook.addWorksheet(
      "Результаты",
      {
        views: [
          {
            state:
              "frozen",

            ySplit:
              1
          }
        ]
      }
    );


  worksheet.addRow(
    COLUMN_NAMES
  );


  for (
    const row of rows
  ) {
    worksheet.addRow(row);
  }


  worksheet.autoFilter = {
    from: "A1",

    to:
      `N${Math.max(
        1,
        rows.length + 1
      )}`
  };


  const header =
    worksheet.getRow(1);


  header.font = {
    bold: true
  };


  header.alignment = {
    horizontal:
      "center",

    vertical:
      "middle",

    wrapText:
      true
  };


  header.fill = {
    type:
      "pattern",

    pattern:
      "solid",

    fgColor: {
      argb:
        "FFD9EAF7"
    }
  };


  const widths = [
    22,
    12,
    20,
    20,
    13,
    28,
    38,
    18,
    48,
    40,
    42,
    22,
    32,
    34
  ];


  widths.forEach(
    (width, index) => {

      worksheet.getColumn(
        index + 1
      ).width =
        width;
    }
  );


  worksheet.eachRow(
    (
      row,
      rowNumber
    ) => {

      if (
        rowNumber > 1
      ) {

        row.alignment = {
          vertical:
            "top",

          wrapText:
            true
        };
      }
    }
  );


  return workbook.xlsx.writeBuffer();
}


exports.handler =
  async function handler(event) {

    resetDiagnosticLog();

    const invocationId =
      Math.random()
        .toString(36)
        .slice(2, 10);


    log(
      "FUNCTION START",
      {
        invocation_id:
          invocationId,

        method:
          event.httpMethod,

        node:
          process.version,

        region:
          process.env.AWS_REGION ||
          null
      }
    );


    if (
      event.httpMethod !==
      "POST"
    ) {

      return {
        statusCode: 405,

        headers: {
          "Content-Type":
            "application/json; charset=utf-8"
        },

        body:
          JSON.stringify({
            error:
              "Method not allowed"
          })
      };
    }


    try {

      const body =
        JSON.parse(
          event.body || "{}"
        );


      const names =
        [
          ...new Set(
            (
              body.trade_names ||
              []
            )
              .map(
                (value) =>
                  String(
                    value
                  ).trim()
              )
              .filter(
                Boolean
              )
          )
        ];


      if (!names.length) {

        throw new Error(
          "Укажите минимум одно торговое наименование"
        );
      }


      if (
        !body.date_from ||
        !body.date_to
      ) {

        throw new Error(
          "Укажите обе даты"
        );
      }


      log(
        "Request accepted",
        {
          invocation_id:
            invocationId,

          trade_names:
            names,

          date_from:
            body.date_from,

          date_to:
            body.date_to
        }
      );


      let rows = [];


      for (
        const name of names
      ) {

        const result =
          await loadTradeName(
            name,
            body.date_from,
            body.date_to
          );


        rows.push(
          ...result
        );
      }


      /*
        Полные дубликаты убираем только
        внутри итогового объединённого массива.
      */

      const seen =
        new Set();


      rows =
        rows.filter(
          (row) => {

            const key =
              JSON.stringify(
                row
              );


            if (
              seen.has(key)
            ) {

              return false;
            }


            seen.add(key);

            return true;
          }
        );


      log(
        "Creating workbook",
        {
          rows:
            rows.length
        }
      );


      const buffer =
        await createWorkbook(
          rows
        );


      const filename =
        safeFilename(
          names
        );


      log(
        "FUNCTION SUCCESS",
        {
          invocation_id:
            invocationId,

          filename,

          rows:
            rows.length,

          bytes:
            buffer.length
        }
      );


      return {
        statusCode: 200,

        isBase64Encoded:
          true,

        headers: {

          "Content-Type":
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

          "Content-Disposition":
            "attachment; " +
            "filename*=UTF-8''" +
            encodeURIComponent(
              filename
            ),

          "X-Filename":
            encodeURIComponent(
              filename
            ),

          "Cache-Control":
            "no-store"
        },

        body:
          Buffer
            .from(buffer)
            .toString(
              "base64"
            )
      };


    } catch (error) {
    
      log(
        "FUNCTION ERROR",
        {
          invocation_id:
            invocationId,
    
          name:
            error.name,
    
          message:
            error.message,
    
          stack:
            error.stack
        }
      );
    
    
      return {
    
        statusCode: 500,
    
        headers: {
    
          "Content-Type":
            "application/json; charset=utf-8",
    
          "Cache-Control":
            "no-store"
        },
    
        body:
          JSON.stringify({
    
            error:
              error.message ||
              String(error),
    
            invocation_id:
              invocationId,
    
            diagnostic_log:
              getDiagnosticLog()
          })
      };
    }
};
