// @ts-check
"use strict";

/**
 * ============================================
 *  GAS 基本テンプレ + 単純テストスイート
 *  - 規約: const + arrow / Logger / strict / 1作用1関数
 *  - 入口はファイル末尾（トリガ・実行対象の公開）
 *  - テストは tests 配列に {name, failMessage, check:()=>boolean} を追加
 *  - 個別実行: const testXxx = testByNames(["A","B"]) / const testFoo = testByName("A")
 * ============================================
 */

/** ===== ログ ========================================== */
/** @type {(msg: string, obj?: unknown) => void} */
const logInfo = (msg, obj) =>
  // @ts-ignore
  Logger.log(`[INFO] ${msg}${obj ? " " + JSON.stringify(obj) : ""}`);
/** @type {(msg: string, obj?: unknown) => void} */
const logWarn = (msg, obj) =>
  // @ts-ignore
  Logger.log(`[WARN] ${msg}${obj ? " " + JSON.stringify(obj) : ""}`);
/** @type {(msg: string, obj?: unknown) => void} */
const logError = (msg, obj) =>
  // @ts-ignore
  Logger.log(`[ERROR] ${msg}${obj ? " " + JSON.stringify(obj) : ""}`);

/** ===== 共通実行ラッパ ================================= */
/**
 * 入口から呼ぶ安全実行ラッパ。開始/終了ログと例外ログを一元化。
 * @template T
 * @param {string} name
 * @param {() => T} thunk
 * @returns {T}
 */
const safeMain = (name, thunk) => {
  const t0 = Date.now();
  try {
    logInfo(`${name} start`);
    const out = thunk();
    logInfo(`${name} done`, { ms: Date.now() - t0 });
    return out;
  } catch (err) {
    const e = /** @type {Error} */ (err);
    logError(`${name} failed`, { message: e.message, stack: e.stack });
    throw e;
  }
};

/** ===== テスト型定義 =================================== */
/** @typedef {{ name: string, failMessage: string, check: () => boolean }} TestCase */
/** @typedef {{ name: string, ok: boolean, message: string, ms: number }} TestResult */

/** ===== テストレジストリ =============================== */
/** @type {Array<TestCase>} */
const tests = [];

/** ===== データローダ: 見積もり必要_テンプレート =================== */
const estimateTemplatesTable = {
  tableName: "見積もり必要_テンプレート",
  headers: {
    name: "名前",
    link: "リンク",
  },
};

/** 型メモ */
/** @typedef {any} GTable */
/** @typedef {any} GGridRange */
/** @typedef {{ tableId: string, sheetId: number, sheetTitle: string, range: GGridRange }} TableMeta */

/** @type {Record<string,string>|undefined} */
let _estimateTemplateCache = undefined;

/** ====== テーブル検索ユーティリティ ====== */
/**
 * スプレッドシート内の全テーブルを列挙し、name->meta の辞書を返す。
 * @returns {Record<string, TableMeta>}
 */
const getTablesIndex = () => {
  // @ts-ignore
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  // Advanced Sheets API を使用（サービスを有効化しておくこと）
  // @ts-ignore
  const resp = Sheets.Spreadsheets.get(spreadsheetId, {
    fields: "sheets(properties(sheetId,title),tables(name,tableId,range))",
  });
  /** @type {Record<string, TableMeta>} */
  const out = {};
  const sheets = resp.sheets || [];
  for (const sh of sheets) {
    const props = sh.properties;
    const sheetId = props && props.sheetId != undefined ? props.sheetId : -1;
    const sheetTitle = props && props.title ? props.title : "";
    const tables = sh.tables || [];
    for (const t of tables) {
      out[String(t.name)] = {
        tableId: String(t.tableId),
        sheetId,
        sheetTitle,
        range: /** @type {GGridRange} */ (t.range),
      };
    }
  }
  return out;
};

/**
 * テーブル名からメタデータを取得
 * @param {string} tableName
 * @returns {TableMeta}
 */
const getTableMetaByName = (tableName) => {
  const idx = getTablesIndex();
  const meta = idx[tableName];
  if (!meta) throw new Error(`table not found: ${tableName}`);
  return meta;
};

/**
 * GridRange -> A1 変換
 * @param {GGridRange} gr
 * @param {string} sheetTitle
 */
const gridRangeToA1 = (gr, sheetTitle) => {
  const toColA1 = (zero) => {
    let n = Number(zero) + 1; // 1-based
    let s = "";
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  };
  const sr = (gr.startRowIndex || 0) + 1; // inclusive (1-based)
  const er = gr.endRowIndex || sr; // exclusive -> inclusive: as-is
  const sc = toColA1(gr.startColumnIndex || 0); // inclusive
  const ec = toColA1((gr.endColumnIndex || 1) - 1); // exclusive -> inclusive
  return `${sheetTitle}!${sc}${sr}:${ec}${er}`;
};

/**
 * 見積もり必要_テンプレート（テーブル）を読み込み、{ 名前: リンク } のマップを返す。
 * @returns {Record<string,string>}
 */
const getEstimateTemplateMap = () => {
  if (_estimateTemplateCache) return _estimateTemplateCache;
  const meta = getTableMetaByName(estimateTemplatesTable.tableName);
  const a1 = gridRangeToA1(meta.range, meta.sheetTitle);
  // @ts-ignore
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  // @ts-ignore
  const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
  const values = vr.values || [];
  if (!values.length) {
    throw new Error("テーブルが空です");
  }

  // ヘッダー行検出
  const header = values[0].map((v) => String(v).trim());
  const nameIdx = header.indexOf(estimateTemplatesTable.headers.name);
  const linkIdx = header.indexOf(estimateTemplatesTable.headers.link);
  if (nameIdx === -1 || linkIdx === -1) {
    throw new Error(
      `ヘッダー未検出: 必要な列名「${estimateTemplatesTable.headers.name}」「${estimateTemplatesTable.headers.link}」`
    );
  }

  /** @type {Record<string,string>} */
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const name = String(row[nameIdx] ?? "").trim();
    const link = String(row[linkIdx] ?? "").trim();
    if (!name) {
      continue;
    }
    if (!link) {
      continue;
    }
    if (Object.prototype.hasOwnProperty.call(map, name) && map[name] !== link) {
      logWarn("duplicate key in 見積もり必要_テンプレート", {
        row: i + 1,
        name,
        prev: map[name],
        next: link,
      });
    }
    map[name] = link;
  }
  _estimateTemplateCache = map;
  logInfo("Loaded table 見積もり必要_テンプレート", {
    a1,
    count: Object.keys(map).length,
    tableId: meta.tableId,
  });
  return map;
};

/** @param {string} name @returns {string|undefined} */
const getEstimateLinkByName = (name) =>
  getEstimateTemplateMap()[String(name).trim()];

/** ===== サンプルテスト（削除/置換OK） ================= */
// 成功する例
tests.push({
  name: "sample:true",
  failMessage: "should be true",
  check: () => true,
});

// 足し算の例
tests.push({
  name: "sample:sum",
  failMessage: "1 + 2 should equal 3",
  check: () => 1 + 2 === 3,
});

// 失敗の例（動作確認用）
// tests.push({
//   name: "sample:fail",
//   failMessage: "this test is expected to fail",
//   check: () => false,
// });

// ここからテーブル読み込みのテストを追加
// 観測用: "dummy" -> "dummy link" を期待
// 必要に応じて実データ側にダミー行を用意してください。
tests.push({
  name: "template:dummy",
  failMessage: 'dummy should map to "dummy link"',
  check: () => getEstimateLinkByName("dummy") === "dummy link",
});

// POグループメンバー: 列存在（ローダが投げなければOK）
tests.push({
  name: "po_members:columns",
  failMessage: "ヘッダー「表示名」「メールアドレス」が存在しません",
  check: () => {
    getPoGroupMembers();
    return true;
  },
});

// POグループメンバー: 件数 >= 1（表示名）
tests.push({
  name: "po_members:nonempty:displayNames",
  failMessage: "表示名の件数が 1 未満",
  check: () => getPoDisplayNames().length >= 1,
});

// POグループメンバー: 件数 >= 1（メールアドレス）
tests.push({
  name: "po_members:nonempty:emails",
  failMessage: "メールアドレスの件数が 1 未満",
  check: () => getPoEmails().length >= 1,
});

// 見積もり履歴: 先頭に 1 行追加（指定データ）
tests.push({
  name: "estimate_history:addRow",
  failMessage: "見積もり履歴への行追加に失敗",
  check: () => {
    addEstimateHistoryTopRow({
      date: "2025-08-30",
      midText: "test 中間スプシ",
      midUrl: "https://www.google.com/",
      formText: "test Google Form",
      formUrl: "https://www.google.com/",
      resultText: "test 結果スプシ",
      resultUrl: "https://www.google.com/",
    });
    return true;
  },
});

// 見積もり必要_メンバー: 列存在（ローダが投げなければOK）
tests.push({
  name: "estimate_required_members:columns",
  failMessage: "ヘッダー「表示名」「メールアドレス」「回答要否」が存在しません",
  check: () => {
    getEstimateRequiredMembers();
    return true;
  },
});

/** ===== 追加: POグループメンバー ローダ =================== */
const poGroupMembersTable = {
  tableName: "POグループメンバー",
  headers: {
    displayName: "表示名",
    email: "メールアドレス",
  },
};

/** @type {{ displayNames: string[], emails: string[] }|undefined } */
let _poMembersCache = undefined;

/**
 * POグループメンバー（テーブル）を読み込み、表示名とメールの配列を返す。
 * @returns {{ displayNames: string[], emails: string[] }}
 */
const getPoGroupMembers = () => {
  if (_poMembersCache) {
    return _poMembersCache;
  }
  const meta = getTableMetaByName(poGroupMembersTable.tableName);
  const a1 = gridRangeToA1(meta.range, meta.sheetTitle);
  // @ts-ignore
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  // @ts-ignore
  const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
  const values = vr.values || [];
  if (!values.length) {
    throw new Error("テーブルが空です: POグループメンバー");
  }

  const header = values[0].map((v) => String(v).trim());
  const dnIdx = header.indexOf(poGroupMembersTable.headers.displayName);
  const emIdx = header.indexOf(poGroupMembersTable.headers.email);
  if (dnIdx === -1 || emIdx === -1) {
    throw new Error(
      `ヘッダー未検出: 必要な列名「${poGroupMembersTable.headers.displayName}」「${poGroupMembersTable.headers.email}」`
    );
  }

  /** @type {string[]} */
  const displayNames = [];
  /** @type {string[]} */
  const emails = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const dn = String(row[dnIdx] ?? "").trim();
    const em = String(row[emIdx] ?? "").trim();
    if (!dn && !em) {
      continue;
    }
    if (dn) {
      displayNames.push(dn);
    }
    if (em) {
      emails.push(em);
    }
  }
  _poMembersCache = { displayNames, emails };
  logInfo("Loaded table POグループメンバー", {
    a1,
    countRows: values.length - 1,
    displayNames: displayNames.length,
    emails: emails.length,
    tableId: meta.tableId,
  });
  return _poMembersCache;
};

/** @returns {string[]} */
const getPoDisplayNames = () => getPoGroupMembers().displayNames;
/** @returns {string[]} */
const getPoEmails = () => getPoGroupMembers().emails;

/** ===== 追加: 見積もり必要_メンバー ローダ =================== */
const estimateRequiredMembersTable = {
  tableName: "見積もり必要_メンバー",
  headers: {
    displayName: "表示名",
    email: "メールアドレス",
    responseRequired: "回答要否",
  },
};

/** @typedef {{ displayName: string, email: string, responseRequired: "不要" | "必要" }} EstimateRequiredMemberRow */
/** @type {Array<EstimateRequiredMemberRow>|undefined} */
let _estimateRequiredMembersCache = undefined;

/**
 * 見積もり必要_メンバー（テーブル）を読み込み、行配列を返す。
 * @returns {Array<EstimateRequiredMemberRow>}
 */
const getEstimateRequiredMembers = () => {
  if (_estimateRequiredMembersCache) return _estimateRequiredMembersCache;
  const meta = getTableMetaByName(estimateRequiredMembersTable.tableName);
  const a1 = gridRangeToA1(meta.range, meta.sheetTitle);
  // @ts-ignore
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  // @ts-ignore
  const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
  const values = vr.values || [];
  if (!values.length) {
    throw new Error("テーブルが空です: 見積もり必要_メンバー");
  }

  const header = values[0].map((v) => String(v).trim());
  const displayNameIdx = header.indexOf(estimateRequiredMembersTable.headers.displayName);
  const emailIdx = header.indexOf(estimateRequiredMembersTable.headers.email);
  const responseRequiredIdx = header.indexOf(estimateRequiredMembersTable.headers.responseRequired);
  if (displayNameIdx === -1 || emailIdx === -1 || responseRequiredIdx === -1) {
    throw new Error(
      `ヘッダー未検出: 必要な列名「${estimateRequiredMembersTable.headers.displayName}」「${estimateRequiredMembersTable.headers.email}」「${estimateRequiredMembersTable.headers.responseRequired}」`
    );
  }

  /** @type {Array<EstimateRequiredMemberRow>} */
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const displayName = String(row[displayNameIdx] ?? "").trim();
    const email = String(row[emailIdx] ?? "").trim();
    const responseRequired = String(row[responseRequiredIdx] ?? "").trim();
    
    if (!displayName && !email) {
      continue;
    }
    
    // 回答要否の値チェック
    if (responseRequired !== "不要" && responseRequired !== "必要") {
      logWarn("invalid responseRequired value in 見積もり必要_メンバー", {
        row: i + 1,
        displayName,
        email,
        responseRequired,
      });
      continue;
    }
    
    rows.push({
      displayName,
      email,
      responseRequired: /** @type {"不要" | "必要"} */ (responseRequired),
    });
  }

  _estimateRequiredMembersCache = rows;
  logInfo("Loaded table 見積もり必要_メンバー", {
    a1,
    countRows: rows.length,
    tableId: meta.tableId,
  });
  return rows;
};

/** ===== 追加: 見積もり履歴（テーブル） 書き込み =================== */
const estimateHistoryTable = {
  tableName: "見積もり履歴",
  headers: {
    date: "見積もり日",
    mid: "中間スプシ",
    form: "Google Form",
    result: "結果スプシ",
  },
};

/**
 * セルにリンクを設定する（表示文字列は値として、リンクは cell に付与）。
 * @param {any} cell
 * @param {string} text
 * @param {string} url
 */
const setCellLink = (cell, text, url) => {
  cell.setValue(text);
  cell.setRichTextValue(
    // @ts-ignore
    SpreadsheetApp.newRichTextValue()
      .setText(text)
      .setLinkUrl(url)
      .build()
  );
};

/**
 * テーブルの「データ先頭」（ヘッダー直下）に 1 行挿入し、値を書き込む。
 * - headerRowCount は 1 と仮定（現行UIの標準）
 * @param {{ date: string, midText: string, midUrl: string, formText: string, formUrl: string, resultText: string, resultUrl: string }} row
 */
const addEstimateHistoryTopRow = (row) => {
  const meta = getTableMetaByName(estimateHistoryTable.tableName);
  const gr = meta.range; // 0-based, end* は exclusive
  const sheetId = meta.sheetId;
  const sheetTitle = meta.sheetTitle;

  const headerRow1 = (gr.startRowIndex || 0) + 1; // 1-based
  const dataTop0 = (gr.startRowIndex || 0) + 1; // ヘッダー1行前提 → データ先頭(0-based)
  const startCol0 = gr.startColumnIndex || 0;
  const endCol0 = gr.endColumnIndex || startCol0 + 1;
  const colCount = endCol0 - startCol0;

  // ヘッダー取得（列位置を名前で合わせる）
  const headerA1 = gridRangeToA1(
    {
      sheetId,
      startRowIndex: headerRow1 - 1,
      endRowIndex: headerRow1,
      startColumnIndex: startCol0,
      endColumnIndex: endCol0,
    },
    sheetTitle
  );
  // @ts-ignore
  const headerVals = (Sheets.Spreadsheets.Values.get(
    // @ts-ignore
    SpreadsheetApp.getActiveSpreadsheet().getId(),
    headerA1
  ).values || [[]])[0].map((v) => String(v).trim());
  const idxByName = (name) => {
    const idx = headerVals.indexOf(name);
    if (idx === -1) {
      throw new Error(`ヘッダー未検出: ${name}`);
    }
    return idx;
  };

  const idxDate = idxByName(estimateHistoryTable.headers.date);
  const idxMid = idxByName(estimateHistoryTable.headers.mid);
  const idxForm = idxByName(estimateHistoryTable.headers.form);
  const idxResult = idxByName(estimateHistoryTable.headers.result);

  /** @type {string[]} */
  const valuesRow = Array(colCount).fill("");
  valuesRow[idxDate] = row.date; // USER_ENTERED → 日付認識
  // ここでは表示文字列のみを書き込み、リンクは後段でセル自体に付与する
  valuesRow[idxMid] = row.midText;
  valuesRow[idxForm] = row.formText;
  valuesRow[idxResult] = row.resultText;

  // 1) データ先頭に 1 行分のスペースを挿入（テーブル幅に限定）
  // 2) 直後にその行に値を書き込む
  const insertRange = {
    sheetId,
    startRowIndex: dataTop0,
    endRowIndex: dataTop0 + 1,
    startColumnIndex: startCol0,
    endColumnIndex: endCol0,
  };

  // @ts-ignore
  Sheets.Spreadsheets.batchUpdate(
    {
      requests: [
        { insertRange: { range: insertRange, shiftDimension: "ROWS" } },
      ],
    },
    // @ts-ignore
    SpreadsheetApp.getActiveSpreadsheet().getId()
  );

  // 値の書き込み（表示文字列のみ）
  const rowA1 = gridRangeToA1(
    {
      sheetId,
      startRowIndex: dataTop0,
      endRowIndex: dataTop0 + 1,
      startColumnIndex: startCol0,
      endColumnIndex: endCol0,
    },
    sheetTitle
  );
  // @ts-ignore
  Sheets.Spreadsheets.Values.update(
    { values: [valuesRow] },
    // @ts-ignore
    SpreadsheetApp.getActiveSpreadsheet().getId(),
    rowA1,
    { valueInputOption: "USER_ENTERED" }
  );

  // 3) セル自体にリンクを付与（HYPERLINK 式は使わない）
  const buildLinkCell = (text, url) => ({
    userEnteredValue: { stringValue: text },
    textFormatRuns: [
      { startIndex: 0, format: { link: { uri: url } } },
    ],
  });

  const linkReq = {
    requests: [
      {
        updateCells: {
          range: {
            sheetId,
            startRowIndex: dataTop0,
            endRowIndex: dataTop0 + 1,
            startColumnIndex: startCol0 + idxMid,
            endColumnIndex: startCol0 + idxMid + 1,
          },
          rows: [ { values: [ buildLinkCell(row.midText, row.midUrl) ] } ],
          fields: "userEnteredValue,textFormatRuns",
        },
      },
      {
        updateCells: {
          range: {
            sheetId,
            startRowIndex: dataTop0,
            endRowIndex: dataTop0 + 1,
            startColumnIndex: startCol0 + idxForm,
            endColumnIndex: startCol0 + idxForm + 1,
          },
          rows: [ { values: [ buildLinkCell(row.formText, row.formUrl) ] } ],
          fields: "userEnteredValue,textFormatRuns",
        },
      },
      {
        updateCells: {
          range: {
            sheetId,
            startRowIndex: dataTop0,
            endRowIndex: dataTop0 + 1,
            startColumnIndex: startCol0 + idxResult,
            endColumnIndex: startCol0 + idxResult + 1,
          },
          rows: [ { values: [ buildLinkCell(row.resultText, row.resultUrl) ] } ],
          fields: "userEnteredValue,textFormatRuns",
        },
      },
    ],
  };

  // @ts-ignore
  Sheets.Spreadsheets.batchUpdate(linkReq, SpreadsheetApp.getActiveSpreadsheet().getId());

  logInfo("addEstimateHistoryTopRow done (links)", { rowA1 });
};

/** ===== 追加: 見積もり必要_締切 ローダ =================== */
const estimateDeadlineTable = {
  tableName: "見積もり必要_締切",
  headers: {
    dueDate: "締切日",
  },
};

/** @typedef {{ dueDate: string }} EstimateDeadlineRow */
/** @type {Array<EstimateDeadlineRow>|undefined} */
let _estimateDeadlineCache = undefined;

/**
 * 見積もり必要_締切（テーブル）を読み込み、行配列を返す（常に長さ1を想定）
 * @returns {Array<EstimateDeadlineRow>}
 */
const getEstimateDeadlines = () => {
  if (_estimateDeadlineCache) return _estimateDeadlineCache;
  const meta = getTableMetaByName(estimateDeadlineTable.tableName);
  const a1 = gridRangeToA1(meta.range, meta.sheetTitle);
  // @ts-ignore
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  // @ts-ignore
  const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
  const values = vr.values || [];
  if (!values.length) {
    throw new Error("テーブルが空です: 見積もり必要_締切");
  }

  const header = values[0].map((v) => String(v).trim());
  const dueIdx = header.indexOf(estimateDeadlineTable.headers.dueDate);
  if (dueIdx === -1) {
    throw new Error(`ヘッダー未検出: 必要な列名「${estimateDeadlineTable.headers.dueDate}」`);
  }

  /** @type {Array<EstimateDeadlineRow>} */
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const due = String(row[dueIdx] ?? "").trim();
    rows.push({ dueDate: due });
  }

  _estimateDeadlineCache = rows;
  logInfo("Loaded table 見積もり必要_締切", {
    a1,
    countRows: rows.length,
    tableId: meta.tableId,
  });
  return rows;
};

const getEstimateDeadline = () => {
  const deadline = getEstimateDeadlines()[0];
  if (!deadline) {
    throw new Error("締切日が設定されていません");
  }
  return deadline;
};

// ===== テスト追加: 見積もり必要_締切 =====
// 列が存在すること
tests.push({
  name: "estimate_deadline:columns",
  failMessage: "ヘッダー『締切日』が存在しません",
  check: () => {
    getEstimateDeadlines();
    return true;
  },
});

// 行数が常に 1 であること
tests.push({
  name: "estimate_deadline:length1",
  failMessage: "行数が 1 ではありません",
  check: () => getEstimateDeadlines().length === 1,
});

/** ===== テストランナー ================================= */
/** @typedef {{ names?: string[] | undefined }} RunTestsCoreInput */

/** @param {RunTestsCoreInput} [input] 指定時はその名前だけ実行 */
const runTestsCore = (input) => {
  const names = input?.names;
  /** @type {Array<TestCase>} */
  const selected =
    Array.isArray(names) && names.length
      ? tests.filter((t) => names.includes(t.name))
      : tests.slice();

  if (Array.isArray(names) && names.length) {
    const known = new Set(selected.map((t) => t.name));
    const unknown = names.filter((n) => !known.has(n));
    if (unknown.length) {
      logWarn("unknown test names", { unknown });
    }
  }

  /** @type {Array<TestResult>} */
  const results = [];

  for (const tc of selected) {
    const t0 = Date.now();
    try {
      const ok = tc.check() === true;
      const ms = Date.now() - t0;
      results.push({
        name: tc.name,
        ok,
        message: ok ? "" : tc.failMessage,
        ms,
      });
      logInfo(`[TEST] ${ok ? "PASS" : "FAIL"} - ${tc.name}`, {
        ms,
        message: ok ? undefined : tc.failMessage,
      });
    } catch (err) {
      const e = /** @type {Error} */ (err);
      const ms = Date.now() - t0;
      results.push({
        name: tc.name,
        ok: false,
        message: `${tc.failMessage} :: ${e.message}`,
        ms,
      });
      logError(`[TEST] EXCEPTION - ${tc.name}`, { ms, error: e.message });
    }
  }

  const pass = results.filter((r) => r.ok).length;
  const fail = results.length - pass;
  logInfo("[TEST] summary", { total: results.length, pass, fail });

  return results;
};

/** ===== 便利関数 ======================================= */

/** 個別テスト実行 */
const runTestByName = (name) =>
  safeMain("runTestByName", () => runTestsCore({ names: [name] }));

/** 指定名の複数テストを実行 */
const runTestsByNames = (names) =>
  safeMain("runTestsByNames", () => runTestsCore({ names }));

/** 個別テストの例（必要に応じて増やす/編集する） */
const testSampleTrue = () => runTestByName("sample:true");
const testSampleSum = () => runTestByName("sample:sum");

/** 複数まとめて実行する例 */
const testSampleCore = () => runTestsByNames(["sample:true", "sample:sum"]);

/** 見積もり必要_テンプレート: ダミー検証 */
const testTemplateDummy = () => runTestByName("template:dummy");

/** POグループメンバー: 列存在・非空 検証 */
const testPoMembersColumns = () => runTestByName("po_members:columns");
const testPoMembersNamesNonEmpty = () =>
  runTestByName("po_members:nonempty:displayNames");
const testPoMembersEmailsNonEmpty = () =>
  runTestByName("po_members:nonempty:emails");

/** まとめて */
const testPoMembersCore = () =>
  runTestsByNames([
    "po_members:columns",
    "po_members:nonempty:displayNames",
    "po_members:nonempty:emails",
  ]);

/** 見積もり履歴: 行追加テスト */
// 見積もり必要_締切: テスト実行ヘルパ
const testEstimateDeadlineColumns = () => runTestByName("estimate_deadline:columns");
const testEstimateDeadlineLength1 = () => runTestByName("estimate_deadline:length1");
const testEstimateDeadlineCore = () =>
  runTestsByNames(["estimate_deadline:columns", "estimate_deadline:length1"]);

/** 見積もり履歴: 行追加テスト */
const testEstimateHistoryAddRow = () =>
  runTestByName("estimate_history:addRow");

/** 見積もり必要_メンバー: テスト実行ヘルパ */
const testEstimateRequiredMembersColumns = () => runTestByName("estimate_required_members:columns");

/** コアテスト（書き込み等の副作用なし）*/
const testCore = () =>
  runTestsByNames([
    "sample:true",
    "sample:sum",
    "template:dummy",
    "po_members:columns",
    "po_members:nonempty:displayNames",
    "po_members:nonempty:emails",
    "estimate_deadline:columns",
    "estimate_deadline:length1",
    "estimate_required_members:columns",
  ]);

/**
 * 追加の個別バンドル例:
 * const testCore = testByNames(["A","B","C"]);
 */
