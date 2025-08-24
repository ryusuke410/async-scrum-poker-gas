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
/** @typedef {{ googleForm: string, midSpreadsheet: string, resultSpreadsheet: string }} EstimateTemplateLinks */

/** @type {EstimateTemplateLinks|undefined} */
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
 * 見積もり必要_テンプレート（テーブル）を読み込み、固定キーのオブジェクトを返す。
 * @returns {EstimateTemplateLinks}
 */
const getEstimateTemplateLinks = () => {
  if (_estimateTemplateCache) {
    return _estimateTemplateCache;
  }
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
  const tempMap = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const name = String(row[nameIdx] ?? "").trim();
    const link = String(row[linkIdx] ?? "").trim();
    if (!name) {
      continue;
    }
    if (!link) {
      logWarn("empty link in 見積もり必要_テンプレート", {
        row: i + 1,
        name,
      });
      continue;
    }
    if (
      Object.prototype.hasOwnProperty.call(tempMap, name) &&
      tempMap[name] !== link
    ) {
      logWarn("duplicate key in 見積もり必要_テンプレート", {
        row: i + 1,
        name,
        prev: tempMap[name],
        next: link,
      });
    }
    tempMap[name] = link;
  }

  // 必要な3つのキーが存在することを確認
  const googleForm = tempMap["Google Form"];
  const midSpreadsheet = tempMap["中間スプシ"];
  const resultSpreadsheet = tempMap["結果スプシ"];

  if (!googleForm) {
    throw new Error("必須項目「Google Form」のリンクが見つかりません");
  }
  if (!midSpreadsheet) {
    throw new Error("必須項目「中間スプシ」のリンクが見つかりません");
  }
  if (!resultSpreadsheet) {
    throw new Error("必須項目「結果スプシ」のリンクが見つかりません");
  }

  /** @type {EstimateTemplateLinks} */
  const links = {
    googleForm,
    midSpreadsheet,
    resultSpreadsheet,
  };

  _estimateTemplateCache = links;
  logInfo("Loaded table 見積もり必要_テンプレート", {
    a1,
    googleForm,
    midSpreadsheet,
    resultSpreadsheet,
    tableId: meta.tableId,
  });
  return links;
};

/** 個別のリンクアクセサ */
const getGoogleFormLink = () => getEstimateTemplateLinks().googleForm;
const getMidSpreadsheetLink = () => getEstimateTemplateLinks().midSpreadsheet;
const getResultSpreadsheetLink = () =>
  getEstimateTemplateLinks().resultSpreadsheet;

/** ===== テンプレートコピー機能 =================== */

/**
 * スプレッドシートをコピーして新しいタイトルを設定
 * @param {string} templateUrl - コピー元のスプレッドシートURL
 * @param {string} newTitle - 新しいスプレッドシートのタイトル
 * @returns {string} - 新しいスプレッドシートのURL
 */
const copySpreadsheetFromUrl = (templateUrl, newTitle) => {
  // URLからスプレッドシートIDを抽出
  const match = templateUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) {
    throw new Error(`Invalid spreadsheet URL: ${templateUrl}`);
  }
  const templateId = match[1];

  // @ts-ignore
  const templateFile = DriveApp.getFileById(templateId);
  // @ts-ignore
  const copiedFile = templateFile.makeCopy(newTitle);
  const copiedId = copiedFile.getId();

  return `https://docs.google.com/spreadsheets/d/${copiedId}/edit`;
};

/**
 * Google FormのURLからコピーを作成
 * @param {string} templateUrl - コピー元のGoogle FormのURL
 * @param {string} newTitle - 新しいGoogle Formのタイトル
 * @returns {string} - 新しいGoogle FormのURL
 */
const copyFormFromUrl = (templateUrl, newTitle) => {
  // URLからFormIDを抽出
  const match = templateUrl.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) {
    throw new Error(`Invalid form URL: ${templateUrl}`);
  }
  const templateId = match[1];

  // @ts-ignore
  const templateFile = DriveApp.getFileById(templateId);
  // @ts-ignore
  const copiedFile = templateFile.makeCopy(newTitle);
  const copiedId = copiedFile.getId();

  return `https://docs.google.com/forms/d/${copiedId}/edit`;
};

/**
 * Google Formとスプレッドシートをリンクする（Formの送信先をSpreadsheetに設定）
 * @param {string} formUrl - Google FormのURL
 * @param {string} spreadsheetUrl - 送信先スプレッドシートのURL
 */
const linkFormToSpreadsheet = (formUrl, spreadsheetUrl) => {
  // FormのURLからIDを抽出
  const formMatch = formUrl.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
  if (!formMatch) {
    throw new Error(`Invalid Google Form URL: ${formUrl}`);
  }
  const formId = formMatch[1];

  // SpreadsheetのURLからIDを抽出
  const spreadsheetMatch = spreadsheetUrl.match(
    /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/
  );
  if (!spreadsheetMatch) {
    throw new Error(`Invalid Spreadsheet URL: ${spreadsheetUrl}`);
  }
  const spreadsheetId = spreadsheetMatch[1];

  try {
    // @ts-ignore
    const form = FormApp.openById(formId);
    // @ts-ignore
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

    // フォームの送信先をスプレッドシートに設定
    // @ts-ignore
    form.setDestination(
      // @ts-ignore
      FormApp.DestinationType.SPREADSHEET,
      spreadsheet.getId()
    );

    logInfo("Form linked to spreadsheet successfully", {
      formId,
      spreadsheetId,
    });
  } catch (err) {
    const e = /** @type {Error} */ (err);
    logError("Failed to link form to spreadsheet", {
      formId,
      spreadsheetId,
      error: e.message,
    });
    throw new Error(`FormとSpreadsheetのリンクに失敗しました: ${e.message}`);
  }
};

/**
 * テンプレートから3つのファイルをコピーして見積もり履歴に追加
 * 締切日を使用してタイトルプレフィックスを生成
 * @param {string} deadlineDate - 締切日（YYYY-MM-DD形式）
 */
const createEstimateFromTemplates = (deadlineDate) => {
  logInfo("createEstimateFromTemplates start");

  // タイトルプレフィックスを生成（締切日 + "async ポーカー"）
  const titlePrefix = `${deadlineDate} async ポーカー`;

  logInfo("Using deadline date for title", { deadlineDate, titlePrefix });

  // テンプレートリンクを取得
  const templates = getEstimateTemplateLinks();

  // 各ファイルをコピー
  const midUrl = copySpreadsheetFromUrl(templates.midSpreadsheet, titlePrefix);
  const formUrl = copyFormFromUrl(templates.googleForm, titlePrefix);
  const resultUrl = copySpreadsheetFromUrl(
    templates.resultSpreadsheet,
    `${titlePrefix}結果`
  );

  logInfo("Files copied successfully", {
    midUrl,
    formUrl,
    resultUrl,
  });

  // Google FormとSpreadsheetをリンク（FormのdestinationをSpreadsheetに設定）
  linkFormToSpreadsheet(formUrl, midUrl);

  logInfo("Form linked to spreadsheet", {
    formUrl,
    midUrl,
  });

  // Form_Responses テーブルにダミー行を追加
  addFormResponsesDummyRow(midUrl);

  logInfo("Added dummy row to Form_Responses table", {
    midUrl,
  });

  // 中間スプシの「メンバー」テーブルを「見積もり必要_メンバー」テーブルのデータで更新
  updateMembersTable(midUrl);

  logInfo("Updated Members table with estimate required members data", {
    midUrl,
  });

  // 見積もり履歴テーブルに行を追加
  addEstimateHistoryTopRow({
    date: deadlineDate,
    midText: titlePrefix,
    midUrl: midUrl,
    formText: titlePrefix,
    formUrl: formUrl,
    resultText: `${titlePrefix}結果`,
    resultUrl: resultUrl,
  });

  logInfo("createEstimateFromTemplates completed", {
    date: deadlineDate,
    titlePrefix,
    deadlineDate,
  });

  return {
    date: deadlineDate,
    titlePrefix,
    midUrl,
    formUrl,
    resultUrl,
  };
};

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
// テンプレートテーブル: 必要な3つのキーが存在し、リンクが空でないこと
tests.push({
  name: "template:required_keys",
  failMessage:
    "Google Form、中間スプシ、結果スプシのいずれかが欠けているか空です",
  check: () => {
    const templates = getEstimateTemplateLinks();
    return !!(
      templates.googleForm &&
      templates.midSpreadsheet &&
      templates.resultSpreadsheet
    );
  },
});

// テンプレートテーブル: 個別アクセサのテスト
tests.push({
  name: "template:google_form_link",
  failMessage: "Google Formのリンクが取得できません",
  check: () => getGoogleFormLink().length > 0,
});

tests.push({
  name: "template:mid_spreadsheet_link",
  failMessage: "中間スプシのリンクが取得できません",
  check: () => getMidSpreadsheetLink().length > 0,
});

tests.push({
  name: "template:result_spreadsheet_link",
  failMessage: "結果スプシのリンクが取得できません",
  check: () => getResultSpreadsheetLink().length > 0,
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

// 見積もり必要_課題リスト: 列存在（ローダが投げなければOK）
tests.push({
  name: "estimate_issue_list:columns",
  failMessage: "ヘッダー「タイトル」「URL」が存在しません",
  check: () => {
    getEstimateIssueList();
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
  if (_estimateRequiredMembersCache) {
    return _estimateRequiredMembersCache;
  }
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
  const displayNameIdx = header.indexOf(
    estimateRequiredMembersTable.headers.displayName
  );
  const emailIdx = header.indexOf(estimateRequiredMembersTable.headers.email);
  const responseRequiredIdx = header.indexOf(
    estimateRequiredMembersTable.headers.responseRequired
  );
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

/** ===== 追加: Form_Responses テーブル書き込み =================== */
const formResponsesTable = {
  tableName: "Form_Responses",
  headers: {
    timestamp: "タイムスタンプ",
    email: "メールアドレス",
    premise: "E1. 見積もりの前提、質問",
    estimate: "E1. 見積り値",
  },
};

/**
 * 指定されたスプレッドシートの Form_Responses テーブルに dummy 行を追加
 * @param {string} spreadsheetUrl - 対象スプレッドシートのURL
 */
const addFormResponsesDummyRow = (spreadsheetUrl) => {
  // SpreadsheetのURLからIDを抽出
  const spreadsheetMatch = spreadsheetUrl.match(
    /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/
  );
  if (!spreadsheetMatch) {
    throw new Error(`Invalid spreadsheet URL: ${spreadsheetUrl}`);
  }
  const spreadsheetId = spreadsheetMatch[1];

  logInfo("Adding dummy row to Form_Responses", { spreadsheetId });

  try {
    // 対象スプレッドシートのテーブル一覧を取得
    // @ts-ignore
    const resp = Sheets.Spreadsheets.get(spreadsheetId, {
      fields: "sheets(properties(sheetId,title),tables(name,tableId,range))",
    });

    const sheets = resp.sheets || [];
    let formResponsesMeta = null;

    // Form_Responses テーブルを探す
    for (const sh of sheets) {
      const tables = sh.tables || [];
      for (const tbl of tables) {
        if (tbl.name === formResponsesTable.tableName) {
          formResponsesMeta = {
            tableId: tbl.tableId,
            sheetId: sh.properties.sheetId,
            sheetTitle: sh.properties.title,
            range: tbl.range,
          };
          break;
        }
      }
      if (formResponsesMeta) break;
    }

    if (!formResponsesMeta) {
      throw new Error(
        `Table not found: ${formResponsesTable.tableName} in spreadsheet ${spreadsheetId}`
      );
    }

    logInfo("Found Form_Responses table", formResponsesMeta);

    // テーブルの現在の範囲を取得してヘッダー行を確認
    const a1 = gridRangeToA1(
      formResponsesMeta.range,
      formResponsesMeta.sheetTitle
    );
    // @ts-ignore
    const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
    const values = vr.values || [];

    if (!values.length) {
      throw new Error(`Table is empty: ${formResponsesTable.tableName}`);
    }

    const header = values[0].map((v) => String(v).trim());
    logInfo("Form_Responses headers", { header });

    // 力技: 普通のスプレッドシートとして行追加
    // テーブルの次の行（ヘッダーの直下）にダミーデータを追加
    const tableStartRow = (formResponsesMeta.range.startRowIndex || 0) + 1; // ヘッダーの次の行（0-based）

    // ダミーデータを準備（タイムスタンプ=0, メールアドレス=dummy）
    const dummyRowData = header.map((headerCell) => {
      const headerName = headerCell.trim();
      if (headerName === formResponsesTable.headers.email) {
        return "dummy";
      }
      return ""; // その他の列は空
    });

    // A1形式の範囲を作成（ヘッダーの直下の行）
    const dummyRowA1 = gridRangeToA1(
      {
        sheetId: formResponsesMeta.sheetId,
        startRowIndex: tableStartRow,
        endRowIndex: tableStartRow + 1,
        startColumnIndex: formResponsesMeta.range.startColumnIndex,
        endColumnIndex: formResponsesMeta.range.endColumnIndex,
      },
      formResponsesMeta.sheetTitle
    );

    // @ts-ignore
    Sheets.Spreadsheets.Values.update(
      { values: [dummyRowData] },
      spreadsheetId,
      dummyRowA1,
      { valueInputOption: "USER_ENTERED" }
    );

    // テーブルの範囲を拡張してダミー行を含める（UpdateTableRequest使用）
    const newEndRowIndex = (formResponsesMeta.range.startRowIndex || 0) + 2; // ヘッダー + ダミー行
    // @ts-ignore
    Sheets.Spreadsheets.batchUpdate(
      {
        requests: [
          {
            updateTable: {
              table: {
                tableId: formResponsesMeta.tableId,
                range: {
                  sheetId: formResponsesMeta.sheetId,
                  startRowIndex: formResponsesMeta.range.startRowIndex,
                  endRowIndex: newEndRowIndex,
                  startColumnIndex: formResponsesMeta.range.startColumnIndex,
                  endColumnIndex: formResponsesMeta.range.endColumnIndex,
                },
              },
              fields: "range",
            },
          },
        ],
      },
      spreadsheetId
    );

    logInfo(
      "Successfully added dummy row to Form_Responses table using Values.update",
      {
        tableId: formResponsesMeta.tableId,
        sheetId: formResponsesMeta.sheetId,
        dummyRowA1,
        dummyRowData,
        newTableEndRow: newEndRowIndex,
      }
    );
  } catch (err) {
    logError("Failed to add dummy row to Form_Responses", {
      error: err.toString(),
      spreadsheetUrl,
    });
    throw err;
  }
};

/** ===== 追加: 中間スプシの「メンバー」テーブル書き込み =================== */
const membersTable = {
  tableName: "メンバー",
  headers: {
    displayName: "表示名",
    email: "メールアドレス",
    responseRequired: "回答要否",
    responseStatus: "回答状況",
  },
};

/**
 * 指定されたスプレッドシートの「メンバー」テーブルを元のスプシの「見積もり必要_メンバー」テーブルのデータで更新
 * @param {string} spreadsheetUrl - 対象スプレッドシートのURL
 */
const updateMembersTable = (spreadsheetUrl) => {
  // SpreadsheetのURLからIDを抽出
  const spreadsheetMatch = spreadsheetUrl.match(
    /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/
  );
  if (!spreadsheetMatch) {
    throw new Error(`Invalid spreadsheet URL: ${spreadsheetUrl}`);
  }
  const spreadsheetId = spreadsheetMatch[1];

  logInfo("Updating Members table with data from estimate required members", {
    spreadsheetId,
  });

  // 元のスプシの「見積もり必要_メンバー」テーブルのデータを取得
  const estimateMembers = getEstimateRequiredMembers();
  logInfo("Retrieved estimate required members", {
    count: estimateMembers.length,
  });

  if (estimateMembers.length === 0) {
    logWarn(
      "No estimate required members found, skipping Members table update"
    );
    return;
  }

  try {
    // 対象スプレッドシートのテーブル一覧を取得
    // @ts-ignore
    const resp = Sheets.Spreadsheets.get(spreadsheetId, {
      fields: "sheets(properties(sheetId,title),tables(name,tableId,range))",
    });

    const sheets = resp.sheets || [];
    let membersMeta = null;

    // メンバーテーブルを探す
    for (const sh of sheets) {
      const tables = sh.tables || [];
      for (const tbl of tables) {
        if (tbl.name === membersTable.tableName) {
          membersMeta = {
            tableId: tbl.tableId,
            sheetId: sh.properties.sheetId,
            sheetTitle: sh.properties.title,
            range: tbl.range,
          };
          break;
        }
      }
      if (membersMeta) break;
    }

    if (!membersMeta) {
      throw new Error(
        `Table not found: ${membersTable.tableName} in spreadsheet ${spreadsheetId}`
      );
    }

    logInfo("Found Members table", membersMeta);

    // テーブルの現在の範囲を取得してヘッダー行を確認
    const a1 = gridRangeToA1(membersMeta.range, membersMeta.sheetTitle);
    // @ts-ignore
    const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
    const values = vr.values || [];

    if (!values.length) {
      throw new Error(`Table is empty: ${membersTable.tableName}`);
    }

    const header = values[0].map((v) => String(v).trim());
    logInfo("Members table headers", { header });

    // ヘッダー位置を取得
    const displayNameIdx = header.indexOf(membersTable.headers.displayName);
    const emailIdx = header.indexOf(membersTable.headers.email);
    const responseRequiredIdx = header.indexOf(
      membersTable.headers.responseRequired
    );
    const responseStatusIdx = header.indexOf(
      membersTable.headers.responseStatus
    );

    if (
      displayNameIdx === -1 ||
      emailIdx === -1 ||
      responseRequiredIdx === -1 ||
      responseStatusIdx === -1
    ) {
      throw new Error("Required headers not found in Members table");
    }

    // データ行の開始位置を計算
    const dataStartRow = (membersMeta.range.startRowIndex || 0) + 1; // ヘッダーの次の行（0-based）
    const currentDataRows = values.length - 1; // ヘッダーを除く既存データ行数
    const requiredRows = estimateMembers.length; // 必要な行数

    // 1. 必要な行数だけ挿入
    if (requiredRows > 0) {
      // @ts-ignore
      Sheets.Spreadsheets.batchUpdate(
        {
          requests: [
            {
              insertDimension: {
                range: {
                  sheetId: membersMeta.sheetId,
                  dimension: "ROWS",
                  startIndex: dataStartRow, // 挿入位置（0-based）
                  endIndex: dataStartRow + requiredRows, // 必要行数分
                },
                inheritFromBefore: false,
              },
            },
          ],
        },
        spreadsheetId
      );
      logInfo("Inserted rows", { count: requiredRows, startRow: dataStartRow });
    }

    // 2. 不要な既存データ行を削除（新しく挿入した行より下）
    if (currentDataRows > 0) {
      const deleteStartRow = dataStartRow + requiredRows; // 新規挿入行の次から
      const deleteEndRow = deleteStartRow + currentDataRows; // 既存データ行数分

      // @ts-ignore
      Sheets.Spreadsheets.batchUpdate(
        {
          requests: [
            {
              deleteDimension: {
                range: {
                  sheetId: membersMeta.sheetId,
                  dimension: "ROWS",
                  startIndex: deleteStartRow,
                  endIndex: deleteEndRow,
                },
              },
            },
          ],
        },
        spreadsheetId
      );
      logInfo("Deleted old data rows", {
        startRow: deleteStartRow,
        endRow: deleteEndRow,
      });
    }

    // 3. テーブルの範囲を更新（ヘッダー + 新しいデータ行数）
    const newTableEndRow =
      (membersMeta.range.startRowIndex || 0) + 1 + requiredRows;
    // @ts-ignore
    Sheets.Spreadsheets.batchUpdate(
      {
        requests: [
          {
            updateTable: {
              table: {
                tableId: membersMeta.tableId,
                range: {
                  sheetId: membersMeta.sheetId,
                  startRowIndex: membersMeta.range.startRowIndex,
                  endRowIndex: newTableEndRow,
                  startColumnIndex: membersMeta.range.startColumnIndex,
                  endColumnIndex: membersMeta.range.endColumnIndex,
                },
              },
              fields: "range",
            },
          },
        ],
      },
      spreadsheetId
    );

    // 4. データを挿入（値貼り付け）
    const columnCount =
      (membersMeta.range.endColumnIndex || 0) -
      (membersMeta.range.startColumnIndex || 0);
    const dataRows = estimateMembers.map((member, index) => {
      const row = Array(columnCount).fill("");
      row[displayNameIdx] = member.displayName;
      row[emailIdx] = member.email;
      row[responseRequiredIdx] = member.responseRequired;
      // 1行目（index === 0）の回答状況には式を、それ以外は空
      row[responseStatusIdx] = (() => {
        if (index !== 0) {
          return "";
        }
        return `\
=LET(
  membersNames,  ${membersTable.tableName}[${membersTable.headers.displayName}],
  membersResponseNecessities, ${membersTable.tableName}[${membersTable.headers.responseRequired}],
  membersEmails, ${membersTable.tableName}[${membersTable.headers.email}],
  estimatesEmails, ${formResponsesTable.tableName}[${formResponsesTable.headers.email}],
  assigneeEmails, FILTER(membersEmails, membersResponseNecessities<>"不要"),
  MAP(
    membersEmails,
    LAMBDA(
      email,
      IF(
        ISERROR(MATCH(email, assigneeEmails, 0)),
        "回答不要",
        IF(
          ISERROR(MATCH(email, estimatesEmails, 0)),
          "未回答",
          "回答済み"
        )
      )
    )
  )
)
`;
      })();
      return row;
    });

    // データ範囲のA1記法を作成
    const dataA1 = gridRangeToA1(
      {
        sheetId: membersMeta.sheetId,
        startRowIndex: dataStartRow,
        endRowIndex: dataStartRow + requiredRows,
        startColumnIndex: membersMeta.range.startColumnIndex,
        endColumnIndex: membersMeta.range.endColumnIndex,
      },
      membersMeta.sheetTitle
    );

    // @ts-ignore
    Sheets.Spreadsheets.Values.update(
      { values: dataRows },
      spreadsheetId,
      dataA1,
      { valueInputOption: "USER_ENTERED" }
    );

    logInfo("Successfully updated Members table", {
      tableId: membersMeta.tableId,
      sheetId: membersMeta.sheetId,
      dataRowsInserted: requiredRows,
      newTableEndRow,
      dataA1,
    });
  } catch (err) {
    logError("Failed to update Members table", {
      error: err.toString(),
      spreadsheetUrl,
    });
    throw err;
  }
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
    SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(url).build()
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
    textFormatRuns: [{ startIndex: 0, format: { link: { uri: url } } }],
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
          rows: [{ values: [buildLinkCell(row.midText, row.midUrl)] }],
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
          rows: [{ values: [buildLinkCell(row.formText, row.formUrl)] }],
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
          rows: [{ values: [buildLinkCell(row.resultText, row.resultUrl)] }],
          fields: "userEnteredValue,textFormatRuns",
        },
      },
    ],
  };

  // @ts-ignore
  Sheets.Spreadsheets.batchUpdate(
    linkReq,
    // @ts-ignore
    SpreadsheetApp.getActiveSpreadsheet().getId()
  );

  logInfo("addEstimateHistoryTopRow done (links)", { rowA1 });
};

/** ===== 追加: 見積もり必要_課題リスト ローダ =================== */
const estimateIssueListTable = {
  tableName: "見積もり必要_課題リスト",
  headers: {
    title: "タイトル",
    url: "URL",
  },
};

/** @typedef {{ title: string, url: string }} EstimateIssueRow */
/** @type {Array<EstimateIssueRow>|undefined} */
let _estimateIssueListCache = undefined;

/**
 * 見積もり必要_課題リスト（テーブル）を読み込み、行配列を返す。
 * @returns {Array<EstimateIssueRow>}
 */
const getEstimateIssueList = () => {
  if (_estimateIssueListCache) {
    return _estimateIssueListCache;
  }
  const meta = getTableMetaByName(estimateIssueListTable.tableName);
  const a1 = gridRangeToA1(meta.range, meta.sheetTitle);
  // @ts-ignore
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  // @ts-ignore
  const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
  const values = vr.values || [];
  if (!values.length) {
    throw new Error("テーブルが空です: 見積もり必要_課題リスト");
  }

  const header = values[0].map((v) => String(v).trim());
  const titleIdx = header.indexOf(estimateIssueListTable.headers.title);
  const urlIdx = header.indexOf(estimateIssueListTable.headers.url);
  if (titleIdx === -1 || urlIdx === -1) {
    throw new Error(
      `ヘッダー未検出: 必要な列名「${estimateIssueListTable.headers.title}」「${estimateIssueListTable.headers.url}」`
    );
  }

  /** @type {Array<EstimateIssueRow>} */
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const title = String(row[titleIdx] ?? "").trim();
    const url = String(row[urlIdx] ?? "").trim();

    if (!title && !url) {
      continue;
    }

    rows.push({
      title,
      url,
    });
  }

  _estimateIssueListCache = rows;
  logInfo("Loaded table 見積もり必要_課題リスト", {
    a1,
    countRows: rows.length,
    tableId: meta.tableId,
  });
  return rows;
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
  if (_estimateDeadlineCache) {
    return _estimateDeadlineCache;
  }
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
    throw new Error(
      `ヘッダー未検出: 必要な列名「${estimateDeadlineTable.headers.dueDate}」`
    );
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

/** 見積もり必要_テンプレート: 固定キーのテスト */
const testTemplateRequiredKeys = () => runTestByName("template:required_keys");
const testTemplateGoogleFormLink = () =>
  runTestByName("template:google_form_link");
const testTemplateMidSpreadsheetLink = () =>
  runTestByName("template:mid_spreadsheet_link");
const testTemplateResultSpreadsheetLink = () =>
  runTestByName("template:result_spreadsheet_link");

/** 見積もり必要_テンプレート: まとめて */
const testTemplateCore = () =>
  runTestsByNames([
    "template:required_keys",
    "template:google_form_link",
    "template:mid_spreadsheet_link",
    "template:result_spreadsheet_link",
  ]);

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
const testEstimateDeadlineColumns = () =>
  runTestByName("estimate_deadline:columns");
const testEstimateDeadlineLength1 = () =>
  runTestByName("estimate_deadline:length1");
const testEstimateDeadlineCore = () =>
  runTestsByNames(["estimate_deadline:columns", "estimate_deadline:length1"]);

/** 見積もり履歴: 行追加テスト */
const testEstimateHistoryAddRow = () =>
  runTestByName("estimate_history:addRow");

/** 見積もり必要_メンバー: テスト実行ヘルパ */
const testEstimateRequiredMembersColumns = () =>
  runTestByName("estimate_required_members:columns");

/** 見積もり必要_課題リスト: テスト実行ヘルパ */
const testEstimateIssueListColumns = () =>
  runTestByName("estimate_issue_list:columns");

/** コアテスト（書き込み等の副作用なし）*/
const testCore = () =>
  runTestsByNames([
    "sample:true",
    "sample:sum",
    "template:required_keys",
    "template:google_form_link",
    "template:mid_spreadsheet_link",
    "template:result_spreadsheet_link",
    "po_members:columns",
    "po_members:nonempty:displayNames",
    "po_members:nonempty:emails",
    "estimate_deadline:columns",
    "estimate_deadline:length1",
    "estimate_required_members:columns",
    "estimate_issue_list:columns",
  ]);

/**
 * 追加の個別バンドル例:
 * const testCore = testByNames(["A","B","C"]);
 */

/** ===== Google Form セクション操作機能 =================== */

/**
 * Google FormのURLからFormオブジェクトを取得
 * @param {string} formUrl - Google FormのURL
 * @returns {any} - Formオブジェクト
 */
const getFormFromUrl = (formUrl) => {
  const match = formUrl.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) {
    throw new Error(`Invalid Google Form URL: ${formUrl}`);
  }
  const formId = match[1];
  // @ts-ignore
  return FormApp.openById(formId);
};

/**
 * Formのタイトルと見積もり課題セクションを課題リストに基づいてセットアップ
 * @param {string} formUrl - Google FormのURL
 * @param {string} title - フォームのタイトル
 * @param {Array<EstimateIssueRow>} issueList - 見積もり課題リスト
 */
const setupFormSections = (formUrl, title, issueList) => {
  const form = getFormFromUrl(formUrl);

  // 1. フォームタイトルを設定
  form.setTitle(title);
  logInfo("Updated form title", { title });

  const targetCount = issueList.length;

  let items = form.getItems();

  logInfo("Initial form structure analysis", {
    totalItems: items.length,
    targetCount,
    issueListCount: issueList.length,
  });

  // 3. PAGE_BREAKを探し、2つ目以降があれば削除
  let pageBreakIndices = [];
  for (let i = 0; i < items.length; i++) {
    // @ts-ignore
    if (items[i].getType() === FormApp.ItemType.PAGE_BREAK) {
      pageBreakIndices.push(i);
    }
  }

  if (pageBreakIndices.length === 0) {
    throw new Error("No PAGE_BREAK found in form");
  }

  // 2つ目以降のPAGE_BREAKとそれ以降のアイテムを削除
  if (pageBreakIndices.length > 1) {
    logInfo("Removing extra PAGE_BREAK items", {
      extraPageBreaks: pageBreakIndices.length - 1,
      firstPageBreakIndex: pageBreakIndices[0],
      itemsToRemove: items.length - pageBreakIndices[1],
    });

    // 後ろから削除（インデックスがずれないように）
    for (let i = items.length - 1; i >= pageBreakIndices[1]; i--) {
      form.deleteItem(i);
    }

    // アイテムリストを再取得
    items = form.getItems();
  }

  const firstPageBreakIndex = pageBreakIndices[0];

  const expectedStructure = [
    // @ts-ignore
    FormApp.ItemType.PAGE_BREAK,
    // @ts-ignore
    FormApp.ItemType.PARAGRAPH_TEXT,
    // @ts-ignore
    FormApp.ItemType.LIST,
  ];

  // 2. PAGE_BREAK後の構造を検証
  if (items.length < firstPageBreakIndex + expectedStructure.length) {
    throw new Error(
      `Expected at least ${
        expectedStructure.length
      } items after PAGE_BREAK, but found ${
        items.length - firstPageBreakIndex - 1
      }`
    );
  }

  for (let i = 0; i < expectedStructure.length; i++) {
    const actualType = items[firstPageBreakIndex + i].getType();
    const expectedType = expectedStructure[i];
    if (actualType !== expectedType) {
      throw new Error(
        `Invalid structure at index ${
          firstPageBreakIndex + i
        }: expected ${expectedType}, but found ${actualType}`
      );
    }
  }

  const expectedTotalItems = firstPageBreakIndex + expectedStructure.length;
  if (items.length !== expectedTotalItems) {
    throw new Error(
      `Expected exactly ${expectedTotalItems} items, but found ${items.length}`
    );
  }

  logInfo("Template structure validated", {
    firstPageBreakIndex,
    templateItems: expectedStructure.length,
  });

  // 3. テンプレートアイテムを取得（PAGE_BREAK, PARAGRAPH_TEXT, LIST）
  const templateSectionHeaderLikeItem = items[firstPageBreakIndex];
  const templatePremiseItem = items[firstPageBreakIndex + 1];
  const templateEstimateItem = items[firstPageBreakIndex + 2];

  // 4. 必要な数だけセットを複製（既に1セットあるので、targetCount - 1 回追加）
  const setsToAdd = targetCount - 1;

  logInfo("Adding estimate section sets", { setsToAdd });

  for (let i = 0; i < setsToAdd; i++) {
    templateSectionHeaderLikeItem.duplicate();
    templatePremiseItem.duplicate();
    templateEstimateItem.duplicate();
  }

  logInfo("Successfully duplicated estimate sections", {
    setsAdded: setsToAdd,
    totalSets: targetCount,
  });

  // アイテムリストを再取得
  items = form.getItems();

  for (let i = 0; i < targetCount; i++) {
    const sectionStartIndex =
      firstPageBreakIndex + i * expectedStructure.length;
    const sectionHeaderLikeItem = items[sectionStartIndex];
    const premiseItem = items[sectionStartIndex + 1];
    const estimateItem = items[sectionStartIndex + 2];

    // 課題リストから対応する課題情報を取得
    const issue = issueList[i];

    sectionHeaderLikeItem.setTitle(issue.url);
    premiseItem.setTitle(`E${i + 1}. 見積もりの前提、質問`);
    estimateItem.setTitle(`E${i + 1}. 見積り値`);

    logInfo(`Updated section ${i + 1} titles`, {
      sectionIndex: i + 1,
      issueTitle: issue.title,
      issueUrl: issue.url,
    });
  }

  logInfo("Successfully updated all section titles with issue URLs", {
    totalSections: targetCount,
  });
};

/** ===== エントリポイント（実行対象の公開） ============= */

/**
 * テンプレートから見積もりファイルセットを作成するエントリポイント
 * 締切日を使用してタイトルプレフィックスを自動生成
 * 使用例: runCreateEstimate()
 */
const runCreateEstimate = () =>
  safeMain("runCreateEstimate", () => {
    const deadline = getEstimateDeadline();
    const deadlineDate = deadline.dueDate;
    createEstimateFromTemplates(deadlineDate);
  });

/**
 * デバッグ用: Google Formのテンプレートコピーと基本セットアップのみ実行
 * 使用例: runDebugFormSetup()
 */
const runDebugFormSetup = () =>
  safeMain("runDebugFormSetup", () => {
    const deadline = getEstimateDeadline();
    const deadlineDate = deadline.dueDate;
    const titlePrefix = `${deadlineDate} async ポーカー`;

    logInfo("Debug: Creating form from template", {
      deadlineDate,
      titlePrefix,
    });

    // テンプレートリンクを取得
    const templates = getEstimateTemplateLinks();

    // Google Formをコピー
    const formUrl = copyFormFromUrl(templates.googleForm, titlePrefix);
    logInfo("Debug: Form copied successfully", { formUrl });

    // 見積もり課題の数を取得
    const issueList = getEstimateIssueList();
    const issueCount = issueList.length;
    logInfo("Debug: Retrieved issue list", { issueCount });

    // フォームのタイトルと見積もり課題セクションをセットアップ
    setupFormSections(formUrl, titlePrefix, issueList);
    logInfo("Debug: Setup form sections completed", {
      title: titlePrefix,
      targetCount: issueCount,
    });

    logInfo("Debug: Form setup completed", {
      formUrl,
      titlePrefix,
      issueCount,
    });

    return {
      formUrl,
      titlePrefix,
      issueCount,
    };
  });
