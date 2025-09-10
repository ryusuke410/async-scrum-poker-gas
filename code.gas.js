/// <reference types="@types/google-apps-script" />
/// <reference types="./types/sheets-advanced" />
/// <reference types="./types/type-tools" />
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
/**
 * @param {string} msg
 * @param {unknown} [obj]
 */
const logInfo = (msg, obj) => {
  Logger.log(`[INFO] ${msg}${obj ? " " + JSON.stringify(obj) : ""}`);
};
/**
 * @param {string} msg
 * @param {unknown} [obj]
 */
const logWarn = (msg, obj) => {
  Logger.log(`[WARN] ${msg}${obj ? " " + JSON.stringify(obj) : ""}`);
};
/**
 * @param {string} msg
 * @param {unknown} [obj]
 */
const logError = (msg, obj) => {
  Logger.log(`[ERROR] ${msg}${obj ? " " + JSON.stringify(obj) : ""}`);
};

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
    const e = err instanceof Error ? err : new Error(String(err));
    logError(`${name} failed`, { message: e.message, stack: e.stack });
    throw e;
  }
};

/** ===== テスト型定義 =================================== */
/** @typedef {{ name: string, failMessage: string, check: () => boolean }} TestCase */
/** @typedef {{ name: string, ok: boolean, message: string, ms: number }} TestResult */

/** ===== テーブル型定義 ================================= */
/** @typedef {{ tableId: string, sheetId: number, sheetTitle: string, range: RequiredToBeDefined<GoogleAppsScript.Sheets.Schema.GridRange> }} TableMeta */

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
/** @typedef {{ googleForm: string, midSpreadsheet: string, resultSpreadsheet: string }} EstimateTemplateLinks */

/** @type {EstimateTemplateLinks|undefined} */
let _estimateTemplateCache = undefined;

/** ====== テーブル検索ユーティリティ ====== */

/** @type {(spreadsheets: GoogleAppsScript.Sheets.Collection.SpreadsheetsCollection | undefined) => spreadsheets is RequiredToBeDefined<GoogleAppsScript.Sheets.Collection.SpreadsheetsCollection>} */
const isSpreadsheetsCollection = (spreadsheets) => {
  return (
    spreadsheets !== undefined &&
    spreadsheets.DeveloperMetadata !== undefined &&
    spreadsheets.Values !== undefined &&
    spreadsheets.Sheets !== undefined
  );
};

/** @type {(range: GoogleAppsScript.Sheets.Schema.GridRange | undefined) => range is RequiredToBeDefined<GoogleAppsScript.Sheets.Schema.GridRange>} */
const isGridRange = (range) => {
  return (
    range !== undefined &&
    range.startRowIndex !== undefined &&
    range.endRowIndex !== undefined &&
    range.startColumnIndex !== undefined &&
    range.endColumnIndex !== undefined
  );
};

/**
 * スプレッドシート内の全テーブルを列挙し、name->meta の辞書を返す。
 * @returns {Record<string, TableMeta>}
 */
const getTablesIndex = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();

  if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
    throw new Error(
      "Advanced Sheets API is not enabled. Please enable it in the Google Apps Script project."
    );
  }

  const resp = Sheets.Spreadsheets.get(spreadsheetId, {
    fields: "sheets(properties(sheetId,title),tables(name,tableId,range))",
  });
  /** @type {Record<string, TableMeta>} */
  const out = {};
  const sheets = resp.sheets || [];
  for (const sh of sheets) {
    const tables = sh.tables || [];
    for (const t of tables) {
      const tableId = t.tableId;
      if (tableId === undefined) {
        throw new Error("Table ID is undefined");
      }
      const sheetId = sh.properties?.sheetId;
      if (sheetId === undefined) {
        throw new Error("Sheet ID is undefined");
      }
      const sheetTitle = sh.properties?.title;
      if (sheetTitle === undefined) {
        throw new Error("Sheet title is undefined");
      }
      const range = t.range;
      if (!isGridRange(range)) {
        throw new Error("Table range is undefined");
      }
      out[String(t.name)] = {
        tableId,
        sheetId,
        sheetTitle,
        range,
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
  if (!meta) {
    throw new Error(`table not found: ${tableName}`);
  }
  return meta;
};

/**
 * GridRange -> A1 変換
 * @param {GoogleAppsScript.Sheets.Schema.GridRange | undefined} gr
 * @param {string} sheetTitle
 */
const gridRangeToA1 = (gr, sheetTitle) => {
  if (!isGridRange(gr)) {
    throw new Error("GridRange or its properties are undefined");
  }
  const toColA1 = /** @param {number} zero */ (zero) => {
    let n = Number(zero) + 1; // 1-based
    let s = "";
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  };
  const sr = gr.startRowIndex + 1; // inclusive (1-based)
  const er = gr.endRowIndex + 1; // exclusive -> inclusive
  const sc = toColA1(gr.startColumnIndex); // inclusive
  const ec = toColA1(gr.endColumnIndex - 1); // exclusive -> inclusive
  return `${sheetTitle}!${sc}${sr}:${ec}${er}`;
};

/**
 * テーブルのヘッダー情報を取得し、列名からインデックスを検索する関数を返す。
 * 既に値配列を持っている場合はそれを利用し、未指定の場合はヘッダー行のみを読み込む。
 * @param {TableMeta} meta
 * @param {string[][]} [values]
 */
const getTableHeaderInfo = (meta, values) => {
  const gr = meta.range;
  if (!gr) {
    throw new Error(`Table range is undefined for ${meta.tableId}`);
  }
  const sheetId = meta.sheetId;
  const sheetTitle = meta.sheetTitle;
  const dataTop0 = (gr.startRowIndex || 0) + 1; // header1 前提 -> データ先頭
  const startCol0 = gr.startColumnIndex || 0;
  const endCol0 = gr.endColumnIndex || startCol0 + 1;

  /** @type {string[]} */
  let headerVals = [];
  if (values && values[0]) {
    headerVals = values[0].map((v) => String(v).trim());
  } else {
    const headerA1 = gridRangeToA1(
      {
        sheetId,
        startRowIndex: dataTop0 - 1,
        endRowIndex: dataTop0,
        startColumnIndex: startCol0,
        endColumnIndex: endCol0,
      },
      sheetTitle
    );
    if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
      throw new Error("Sheets.Spreadsheets is not available");
    }
    headerVals =
      (Sheets.Spreadsheets.Values.get(
        SpreadsheetApp.getActiveSpreadsheet().getId(),
        headerA1
      ).values || [[]])[0]?.map((v) => String(v).trim()) || [];
  }
  const idxByName = /** @param {string} name */ (name) => {
    const idx = headerVals.indexOf(name);
    if (idx === -1) {
      throw new Error(`ヘッダー未検出: ${name}`);
    }
    return idx;
  };
  return {
    sheetId,
    sheetTitle,
    startCol0,
    endCol0,
    dataTop0,
    headerVals,
    idxByName,
  };
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();

  if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
    throw new Error(
      "Advanced Sheets API is not enabled. Please enable it in the Google Apps Script project."
    );
  }
  const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
  const values = vr.values || [];
  if (!values.length || !values[0]) {
    throw new Error("テーブルが空です");
  }

  const { idxByName } = getTableHeaderInfo(meta, values);
  const nameIdx = idxByName(estimateTemplatesTable.headers.name);
  const linkIdx = idxByName(estimateTemplatesTable.headers.link);

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
  const templateId = extractSpreadsheetIdFromUrl(templateUrl);

  const templateFile = DriveApp.getFileById(templateId);
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
  const templateId = extractFormIdFromUrl(templateUrl);

  const templateFile = DriveApp.getFileById(templateId);
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
  const formId = extractFormIdFromUrl(formUrl);

  // SpreadsheetのURLからIDを抽出
  const spreadsheetId = extractSpreadsheetIdFromUrl(spreadsheetUrl);

  try {
    const form = FormApp.openById(formId);
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

    // フォームの送信先をスプレッドシートに設定
    form.setDestination(
      FormApp.DestinationType.SPREADSHEET,
      spreadsheet.getId()
    );

    logInfo("Form linked to spreadsheet successfully", {
      formId,
      spreadsheetId,
    });
  } catch (err) {
    const e = err instanceof Error ? err : new Error(String(err));
    logError("Failed to link form to spreadsheet", {
      formId,
      spreadsheetId,
      error: e.message,
    });
    throw new Error(`FormとSpreadsheetのリンクに失敗しました: ${e.message}`);
  }
};

/**
 * ファイルIDから編集権限をPOグループメンバーに付与する（通知なし）
 * @param {string} fileId - ファイルID
 * @param {string} fileType - ファイルの種類（ログ用）
 */
const grantEditPermissionToPoGroup = (fileId, fileType) => {
  const poEmails = getPoEmails();

  if (!poEmails.length) {
    logWarn(
      `PO group has no email addresses, skipping permission setup for ${fileType}`
    );
    return;
  }

  logInfo(
    `Granting edit permissions to PO group for ${fileType} (no notification)`,
    {
      fileId,
      emailCount: poEmails.length,
    }
  );

  for (const email of poEmails) {
    try {
      const permission = {
        type: "user",
        role: "writer",
        emailAddress: email,
      };

      Drive.Permissions.create(permission, fileId, {
        sendNotificationEmail: false,
      });

      logInfo(
        `Edit permission granted to ${email} for ${fileType} (no notification)`,
        { fileId }
      );
    } catch (err) {
      const e = err instanceof Error ? err : new Error(String(err));
      logWarn(`Failed to grant permission to ${email} for ${fileType}`, {
        fileId,
        error: e.message,
      });
    }
  }
};

/**
 * ファイルIDから閲覧権限を見積もり必要メンバー全員に付与する（通知なし）
 * POメンバーは既に編集権限を持っているため除外する
 * @param {string} fileId - ファイルID
 * @param {string} fileType - ファイルの種類（ログ用）
 */
const grantViewPermissionToEstimateMembers = (fileId, fileType) => {
  const members = getEstimateRequiredMembers();
  const poEmails = getPoEmails();

  // POメンバーを除外した見積もり必要メンバーのメールアドレス一覧を作成
  const memberEmails = members
    .map((member) => member.email)
    .filter((email) => email && !poEmails.includes(email));

  if (!memberEmails.length) {
    logWarn(
      `No email addresses found in estimate required members (excluding PO members), skipping permission setup for ${fileType}`
    );
    return;
  }

  logInfo(
    `Granting view permissions to estimate members for ${fileType} (no notification, excluding PO members)`,
    {
      fileId,
      emailCount: memberEmails.length,
      totalMembersCount: members.length,
      excludedPoCount: members.length - memberEmails.length,
    }
  );

  for (const email of memberEmails) {
    try {
      const permission = {
        type: "user",
        role: "reader",
        emailAddress: email,
      };

      Drive.Permissions.create(permission, fileId, {
        sendNotificationEmail: false,
      });

      logInfo(
        `View permission granted to ${email} for ${fileType} (no notification)`,
        { fileId }
      );
    } catch (err) {
      const e = err instanceof Error ? err : new Error(String(err));
      logWarn(`Failed to grant permission to ${email} for ${fileType}`, {
        fileId,
        error: e.message,
      });
    }
  }
};

/**
 * フォームIDから回答権限を見積もり必要メンバーに付与する
 * POメンバーは既に編集権限を持っているため除外する
 * 見積もりが必要なメンバーには通知を送信し、それ以外には通知を送信しない
 * @param {string} formId - フォームID
 */
const grantFormResponsePermissionToEstimateMembers = (formId) => {
  const members = getEstimateRequiredMembers();
  const poEmails = getPoEmails();

  // POメンバーを除外した見積もり必要メンバーのメールアドレス一覧を作成
  const memberEmails = members
    .map((member) => member.email)
    .filter((email) => email && !poEmails.includes(email));

  if (!memberEmails.length) {
    logWarn(
      `No email addresses found in estimate required members (excluding PO members), skipping form response permission setup`
    );
    return;
  }

  logInfo(
    `Granting form response permissions to estimate members (excluding PO members)`,
    {
      formId,
      emailCount: memberEmails.length,
      totalMembersCount: members.length,
      excludedPoCount: members.length - memberEmails.length,
    }
  );

  for (const email of memberEmails) {
    try {
      // 見積もりが必要なメンバーかどうかを判定
      const member = members.find((m) => m.email === email);
      const needsEstimate = member && member.responseRequired === "必要";

      const permission = {
        role: "reader",
        type: "user",
        emailAddress: email,
        view: "published",
      };

      Drive.Permissions.create(permission, formId, {
        sendNotificationEmail: needsEstimate,
      });

      logInfo(
        `Form response permission granted to ${email} (notification: ${needsEstimate})`,
        { formId, needsEstimate }
      );
    } catch (err) {
      const e = err instanceof Error ? err : new Error(String(err));
      logWarn(`Failed to grant form response permission to ${email}`, {
        formId,
        error: e.message,
      });
    }
  }
};

/**
 * URLからファイルIDを抽出する
 * @param {string} url - Google DriveファイルのURL
 * @param {string} fileType - ファイルの種類（エラーメッセージ用）
 * @returns {string} ファイルID
 */
const extractFileIdFromUrl = (url, fileType) => {
  const fileId = extractFileIdFromUrlOrUndefined(url);
  if (fileId === undefined) {
    throw new Error(`Could not extract file ID from ${fileType} URL: ${url}`);
  }
  return fileId;
};

/**
 * SpreadsheetのURLからファイルIDを抽出する
 * @param {string} url - Google SpreadsheetのURL
 * @returns {string|undefined} ファイルID、見つからない場合はundefined
 */
const extractSpreadsheetIdFromUrlOrUndefined = (url) => {
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match && match[1] ? match[1] : undefined;
};

/**
 * SpreadsheetのURLからファイルIDを抽出する
 * @param {string} url - Google SpreadsheetのURL
 * @returns {string} ファイルID
 * @throws {Error} SpreadsheetのIDが見つからない場合
 */
const extractSpreadsheetIdFromUrl = (url) => {
  const spreadsheetId = extractSpreadsheetIdFromUrlOrUndefined(url);
  if (spreadsheetId === undefined) {
    throw new Error(`Could not extract spreadsheet ID from URL: ${url}`);
  }
  return spreadsheetId;
};

/**
 * Google FormのURLからファイルIDを抽出する
 * @param {string} url - Google FormのURL
 * @returns {string|undefined} ファイルID、見つからない場合はundefined
 */
const extractFormIdFromUrlOrUndefined = (url) => {
  const match = url.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
  return match && match[1] ? match[1] : undefined;
};

/**
 * Google FormのURLからファイルIDを抽出する
 * @param {string} url - Google FormのURL
 * @returns {string} ファイルID
 * @throws {Error} FormのIDが見つからない場合
 */
const extractFormIdFromUrl = (url) => {
  const formId = extractFormIdFromUrlOrUndefined(url);
  if (formId === undefined) {
    throw new Error(`Could not extract form ID from URL: ${url}`);
  }
  return formId;
};

/**
 * Google DriveファイルのURLからファイルIDを抽出する
 * @param {string} url - Google DriveファイルのURL
 * @returns {string|undefined} ファイルID、見つからない場合はundefined
 */
const extractDriveFileIdFromUrlOrUndefined = (url) => {
  const match = url.match(/\/file\/d\/([a-zA-Z0-9-_]+)/);
  return match && match[1] ? match[1] : undefined;
};

/**
 * Google DriveファイルのURLからファイルIDを抽出する
 * @param {string} url - Google DriveファイルのURL
 * @returns {string} ファイルID
 * @throws {Error} DriveファイルのIDが見つからない場合
 */
const extractDriveFileIdFromUrl = (url) => {
  const driveFileId = extractDriveFileIdFromUrlOrUndefined(url);
  if (driveFileId === undefined) {
    throw new Error(`Could not extract drive file ID from URL: ${url}`);
  }
  return driveFileId;
};

/**
 * URLからファイルIDを抽出する（見つからない場合はundefinedを返す）
 * @param {string} url - Google DriveファイルのURL
 * @returns {string|undefined} ファイルID、見つからない場合はundefined
 */
const extractFileIdFromUrlOrUndefined = (url) => {
  // Spreadsheet用のパターン
  const spreadsheetId = extractSpreadsheetIdFromUrlOrUndefined(url);
  if (spreadsheetId) {
    return spreadsheetId;
  }

  // Form用のパターン
  const formId = extractFormIdFromUrlOrUndefined(url);
  if (formId) {
    return formId;
  }

  // 一般的なDriveファイル用のパターン
  const driveFileId = extractDriveFileIdFromUrlOrUndefined(url);
  if (driveFileId) {
    return driveFileId;
  }

  return undefined;
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

  // Google Formのタイトルとセクションをセットアップ
  const issueList = getEstimateIssueList().filter(({ title, url }) => title || url);
  setupFormSections(formUrl, titlePrefix, issueList);

  logInfo("Form sections setup completed", {
    formUrl,
    titlePrefix,
    issueCount: issueList.length,
  });

  // Google Formの回答を可能にする
  const form = getFormFromUrl(formUrl);
  form.setAcceptingResponses(true);

  const formResponseUrl = form.getPublishedUrl();

  logInfo("Form accepting responses enabled", {
    formUrl,
    formResponseUrl,
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

  // 中間スプシの「結果まとめ」テーブルを見積もり課題リストのデータで更新
  updateResultSummaryTable(midUrl);

  logInfo("Updated ResultSummary table with estimate issue list data", {
    midUrl,
  });

  // Slack メッセージを生成
  const members = getEstimateRequiredMembers();
  const mentionList = members
    .filter((m) => m.responseRequired === "必要")
    .map((m) => m.slackMention)
    .join(" ");

  /** @type {RichText} */
  const requestSlackMessage = {
    elements: [
      {
        type: "plain",
        text:
          `${mentionList}\n\n${deadlineDate} の非同期ポーカーです。\n` +
          `締切を${deadlineDate} 16:00 に設定しています。\n\nお手数ですが、`,
      },
      { type: "link", text: "こちら", url: formResponseUrl },
      { type: "plain", text: "からご回答のほどよろしくお願いいたします。" },
    ],
  };

  /** @type {RichText} */
  const completionSlackMessage = {
    elements: [
      { type: "plain", text: "ご回答ありがとうございます。\n\n結果を" },
      { type: "link", text: "こちら", url: resultUrl },
      {
        type: "plain",
        text:
          "にまとめましたので、ご確認のほどよろしくお願いいたします。\n" +
          "特に violation がでた部分については、再見積もりとなりますので、次回の見積もりのためにご参考ください。",
      },
    ],
  };

  // 見積もり履歴テーブルに行を追加
  addEstimateHistoryTopRow({
    date: deadlineDate,
    midText: titlePrefix,
    midUrl: midUrl,
    formText: titlePrefix,
    formUrl: formResponseUrl,
    resultText: `${titlePrefix}結果`,
    resultUrl: resultUrl,
    requestSlackMessage,
    completionSlackMessage,
  });

  try {
    const midFileId = extractSpreadsheetIdFromUrl(midUrl);
    const formFileId = extractFormIdFromUrl(formUrl);
    const resultFileId = extractSpreadsheetIdFromUrl(resultUrl);

    grantEditPermissionToPoGroup(midFileId, "中間スプシ");
    grantEditPermissionToPoGroup(formFileId, "Google Form");
    grantEditPermissionToPoGroup(resultFileId, "結果スプシ");

    logInfo("PO group permissions granted successfully", {
      midUrl,
      resultUrl,
    });

    // フォームに見積もりメンバー全員の回答権限を付与（POメンバー除外、通知は必要な人のみ）
    grantFormResponsePermissionToEstimateMembers(formFileId);
    logInfo("Form response permissions granted to estimate members", {
      formUrl,
    });

    // 結果スプシに見積もりメンバー全員の閲覧権限を付与
    grantViewPermissionToEstimateMembers(resultFileId, "結果スプシ");
    logInfo("Estimate members view permissions granted successfully", {
      resultUrl,
    });
  } catch (err) {
    const e = err instanceof Error ? err : new Error(String(err));
    logWarn("Failed to grant PO group permissions", {
      error: e.message,
      midUrl,
      formUrl,
      resultUrl,
    });
  }

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
      requestSlackMessage: {
        elements: [{ type: "plain", text: "test request" }],
      },
      completionSlackMessage: {
        elements: [{ type: "plain", text: "test completion" }],
      },
    });
    return true;
  },
});

// 見積もり必要_メンバー: 列存在（ローダが投げなければOK）
tests.push({
  name: "estimate_required_members:columns",
  failMessage:
    "ヘッダー「表示名」「メールアドレス」「回答要否」「Slack メンション名」が存在しません",
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
    throw new Error("Sheets.Spreadsheets is not available");
  }
  const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
  const values = vr.values || [];
  if (!values.length || !values[0]) {
    throw new Error("テーブルが空です: POグループメンバー");
  }

  const { idxByName } = getTableHeaderInfo(meta, values);
  const dnIdx = idxByName(poGroupMembersTable.headers.displayName);
  const emIdx = idxByName(poGroupMembersTable.headers.email);

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
    slackMention: "Slack メンション名",
  },
};

/** @typedef {{ displayName: string, email: string, responseRequired: "不要" | "必要", slackMention: string }} EstimateRequiredMemberRow */
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
    throw new Error("Sheets.Spreadsheets is not available");
  }
  const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
  const values = vr.values || [];
  if (!values.length || !values[0]) {
    throw new Error("テーブルが空です: 見積もり必要_メンバー");
  }

  const { idxByName } = getTableHeaderInfo(meta, values);
  const displayNameIdx = idxByName(
    estimateRequiredMembersTable.headers.displayName
  );
  const emailIdx = idxByName(estimateRequiredMembersTable.headers.email);
  const responseRequiredIdx = idxByName(
    estimateRequiredMembersTable.headers.responseRequired
  );
  const slackMentionIdx = idxByName(
    estimateRequiredMembersTable.headers.slackMention
  );

  /** @type {Array<EstimateRequiredMemberRow>} */
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const displayName = String(row[displayNameIdx] ?? "").trim();
    const email = String(row[emailIdx] ?? "").trim();
    const responseRequired = String(row[responseRequiredIdx] ?? "").trim();
    const slackMention = String(row[slackMentionIdx] ?? "").trim();

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
      slackMention,
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
  const spreadsheetId = extractSpreadsheetIdFromUrl(spreadsheetUrl);

  if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
    throw new Error(
      "Advanced Sheets API is not enabled. Please enable it in the Google Apps Script project."
    );
  }

  logInfo("Adding dummy row to Form_Responses", { spreadsheetId });

  try {
    // 対象スプレッドシートのテーブル一覧を取得
    const resp = Sheets.Spreadsheets.get(spreadsheetId, {
      fields: "sheets(properties(sheetId,title),tables(name,tableId,range))",
    });

    const sheets = resp.sheets || [];
    /** @type {TableMeta | undefined} */
    let formResponsesMeta = undefined;

    // Form_Responses テーブルを探す
    for (const sh of sheets) {
      const tables = sh.tables || [];
      for (const tbl of tables) {
        if (tbl.name === formResponsesTable.tableName) {
          const tableId = tbl.tableId;
          if (tableId === undefined) {
            throw new Error("Table ID is undefined");
          }
          const sheetId = sh.properties?.sheetId;
          if (sheetId === undefined) {
            throw new Error("Sheet ID is undefined");
          }
          const sheetTitle = sh.properties?.title;
          if (sheetTitle === undefined) {
            throw new Error("Sheet title is undefined");
          }
          const range = tbl.range;
          if (!isGridRange(range)) {
            throw new Error("Table range is undefined");
          }

          formResponsesMeta = {
            tableId,
            sheetId,
            sheetTitle,
            range,
          };
          break;
        }
      }
      if (formResponsesMeta) {
        break;
      }
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
    if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
      throw new Error("Sheets.Spreadsheets is not available");
    }
    const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
    const values = vr.values || [];

    if (!values.length) {
      throw new Error(`Table is empty: ${formResponsesTable.tableName}`);
    }

    const header = values[0]?.map((v) => String(v).trim());
    if (!header) {
      throw new Error("Header row is undefined");
    }
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
      if (headerName.endsWith(". 見積もりの前提、質問")) {
        return "dummy premise";
      }
      if (headerName.endsWith(". 見積り値")) {
        return "skip";
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

    if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
      throw new Error("Sheets.Spreadsheets is not available");
    }
    Sheets.Spreadsheets.Values.update(
      { values: [dummyRowData] },
      spreadsheetId,
      dummyRowA1,
      { valueInputOption: "USER_ENTERED" }
    );

    // テーブルの範囲を拡張してダミー行を含める（UpdateTableRequest使用）
    const newEndRowIndex = (formResponsesMeta.range.startRowIndex || 0) + 2; // ヘッダー + ダミー行
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
      error: String(err),
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
  const spreadsheetId = extractSpreadsheetIdFromUrl(spreadsheetUrl);

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

  if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
    throw new Error(
      "Advanced Sheets API is not enabled. Please enable it in the Google Apps Script project."
    );
  }

  try {
    // 対象スプレッドシートのテーブル一覧を取得
    const resp = Sheets.Spreadsheets.get(spreadsheetId, {
      fields: "sheets(properties(sheetId,title),tables(name,tableId,range))",
    });

    const sheets = resp.sheets || [];
    /** @type {TableMeta | undefined} */
    let membersMeta = undefined;

    // メンバーテーブルを探す
    for (const sh of sheets) {
      const tables = sh.tables || [];
      for (const tbl of tables) {
        if (tbl.name === membersTable.tableName) {
          const tableId = tbl.tableId;
          if (tableId === undefined) {
            throw new Error("Table ID is undefined");
          }
          const sheetId = sh.properties?.sheetId;
          if (sheetId === undefined) {
            throw new Error("Sheet ID is undefined");
          }
          const sheetTitle = sh.properties?.title;
          if (sheetTitle === undefined) {
            throw new Error("Sheet title is undefined");
          }
          const range = tbl.range;
          if (!isGridRange(range)) {
            throw new Error("Table range is undefined");
          }

          membersMeta = {
            tableId,
            sheetId,
            sheetTitle,
            range,
          };
          break;
        }
      }
      if (membersMeta) {
        break;
      }
    }

    if (!membersMeta) {
      throw new Error(
        `Table not found: ${membersTable.tableName} in spreadsheet ${spreadsheetId}`
      );
    }

    if (!membersMeta.range || !membersMeta.sheetTitle) {
      throw new Error(`Invalid table metadata for ${membersTable.tableName}`);
    }

    logInfo("Found Members table", membersMeta);

    // テーブルの現在の範囲を取得してヘッダー行を確認
    const a1 = gridRangeToA1(membersMeta.range, membersMeta.sheetTitle);
    if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
      throw new Error("Sheets.Spreadsheets is not available");
    }
    const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
    const values = vr.values || [];

    if (!values.length || !values[0]) {
      throw new Error(`Table is empty: ${membersTable.tableName}`);
    }

    const { idxByName, headerVals } = getTableHeaderInfo(membersMeta, values);
    logInfo("Members table headers", { header: headerVals });
    const displayNameIdx = idxByName(membersTable.headers.displayName);
    const emailIdx = idxByName(membersTable.headers.email);
    const responseRequiredIdx = idxByName(
      membersTable.headers.responseRequired
    );
    const responseStatusIdx = idxByName(
      membersTable.headers.responseStatus
    );

    // データ行の開始位置を計算
    const dataStartRow = (membersMeta.range.startRowIndex || 0) + 1; // ヘッダーの次の行（0-based）
    const currentDataRows = values.length - 1; // ヘッダーを除く既存データ行数
    const requiredRows = estimateMembers.length; // 必要な行数

    // 1. 必要な行数だけ挿入
    if (requiredRows > 0) {
      if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
        throw new Error("Sheets.Spreadsheets is not available");
      }
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
  assigneeEmails, FILTER(membersEmails, membersResponseNecessities="必要"),
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
    const error = err instanceof Error ? err : new Error(String(err));
    logError("Failed to update Members table", {
      error: error.toString(),
      spreadsheetUrl,
    });
    throw error;
  }
};

/** ===== 追加: 中間スプシの「結果まとめ」テーブル書き込み =================== */
const resultSummaryTable = {
  tableName: "結果まとめ",
  headers: {
    id: "ID",
    estimateTarget: "見積もり対象",
    status: "ステータス",
    average: "average",
    responseSummary: "回答まとめ",
    min: "min",
    max: "max",
    minBy: "min by",
    maxBy: "max by",
  },
};

/**
 * 指定されたスプレッドシートの「結果まとめ」テーブルを見積もり課題リストのデータで更新
 * @param {string} spreadsheetUrl - 対象スプレッドシートのURL
 */
const updateResultSummaryTable = (spreadsheetUrl) => {
  // SpreadsheetのURLからIDを抽出
  const spreadsheetId = extractSpreadsheetIdFromUrl(spreadsheetUrl);

  logInfo("Updating ResultSummary table with data from estimate issue list", {
    spreadsheetId,
  });

  // 見積もり課題リストを取得
  const issueList = getEstimateIssueList().filter(({ title, url }) => title || url);
  logInfo("Retrieved estimate issue list", {
    count: issueList.length,
  });

  if (issueList.length === 0) {
    logWarn("No estimate issues found, skipping ResultSummary table update");
    return;
  }

  if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
    throw new Error(
      "Advanced Sheets API is not enabled. Please enable it in the Google Apps Script project."
    );
  }

  try {
    // 対象スプレッドシートのテーブル一覧を取得
    const resp = Sheets.Spreadsheets.get(spreadsheetId, {
      fields: "sheets(properties(sheetId,title),tables(name,tableId,range))",
    });

    const sheets = resp.sheets || [];
    /** @type {TableMeta | undefined} */
    let resultSummaryMeta = undefined;

    // 結果まとめテーブルを探す
    for (const sh of sheets) {
      const tables = sh.tables || [];
      for (const tbl of tables) {
        if (tbl.name === resultSummaryTable.tableName) {
          const tableId = tbl.tableId;
          if (tableId === undefined) {
            throw new Error("Table ID is undefined");
          }
          const sheetId = sh.properties?.sheetId;
          if (sheetId === undefined) {
            throw new Error("Sheet ID is undefined");
          }
          const sheetTitle = sh.properties?.title;
          if (sheetTitle === undefined) {
            throw new Error("Sheet title is undefined");
          }
          const range = tbl.range;
          if (!isGridRange(range)) {
            throw new Error("Table range is undefined");
          }
          resultSummaryMeta = {
            tableId,
            sheetId,
            sheetTitle,
            range,
          };
          break;
        }
      }
      if (resultSummaryMeta) {
        break;
      }
    }

    if (!resultSummaryMeta) {
      throw new Error(
        `Table not found: ${resultSummaryTable.tableName} in spreadsheet ${spreadsheetId}`
      );
    }

    if (!resultSummaryMeta.range || !resultSummaryMeta.sheetTitle) {
      throw new Error(
        `Invalid table metadata for ${resultSummaryTable.tableName}`
      );
    }

    logInfo("Found ResultSummary table", resultSummaryMeta);

    // テーブルの現在の範囲を取得してヘッダー行を確認
    const a1 = gridRangeToA1(
      resultSummaryMeta.range,
      resultSummaryMeta.sheetTitle
    );
    const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
    const values = vr.values || [];

    if (!values.length || !values[0]) {
      throw new Error(`Table is empty: ${resultSummaryTable.tableName}`);
    }

    const { idxByName, headerVals } = getTableHeaderInfo(
      resultSummaryMeta,
      values
    );
    logInfo("ResultSummary table headers", { header: headerVals });
    const idIdx = idxByName(resultSummaryTable.headers.id);
    const estimateTargetIdx = idxByName(
      resultSummaryTable.headers.estimateTarget
    );
    const statusIdx = idxByName(resultSummaryTable.headers.status);
    const averageIdx = idxByName(resultSummaryTable.headers.average);
    const responseSummaryIdx = idxByName(
      resultSummaryTable.headers.responseSummary
    );
    const minIdx = idxByName(resultSummaryTable.headers.min);
    const maxIdx = idxByName(resultSummaryTable.headers.max);
    const minByIdx = idxByName(resultSummaryTable.headers.minBy);
    const maxByIdx = idxByName(resultSummaryTable.headers.maxBy);

    // データ行の開始位置を計算
    const dataStartRow = (resultSummaryMeta.range.startRowIndex || 0) + 1; // ヘッダーの次の行（0-based）
    const currentDataRows = values.length - 1; // ヘッダーを除く既存データ行数
    const requiredRows = issueList.length; // 必要な行数

    // 1. 必要な行数だけ挿入
    if (requiredRows > 0) {
      if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
        throw new Error("Sheets.Spreadsheets is not available");
      }
      Sheets.Spreadsheets.batchUpdate(
        {
          requests: [
            {
              insertDimension: {
                range: {
                  sheetId: resultSummaryMeta.sheetId,
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

      if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
        throw new Error("Sheets.Spreadsheets is not available");
      }
      Sheets.Spreadsheets.batchUpdate(
        {
          requests: [
            {
              deleteDimension: {
                range: {
                  sheetId: resultSummaryMeta.sheetId,
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
      (resultSummaryMeta.range.startRowIndex || 0) + 1 + requiredRows;
    if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
      throw new Error("Sheets.Spreadsheets is not available");
    }
    Sheets.Spreadsheets.batchUpdate(
      {
        requests: [
          {
            updateTable: {
              table: {
                tableId: resultSummaryMeta.tableId,
                range: {
                  sheetId: resultSummaryMeta.sheetId,
                  startRowIndex: resultSummaryMeta.range.startRowIndex,
                  endRowIndex: newTableEndRow,
                  startColumnIndex: resultSummaryMeta.range.startColumnIndex,
                  endColumnIndex: resultSummaryMeta.range.endColumnIndex,
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
      (resultSummaryMeta.range.endColumnIndex || 0) -
      (resultSummaryMeta.range.startColumnIndex || 0);
    const dataRows = issueList.map((issue, index) => {
      const row = Array(columnCount).fill("");
      row[idIdx] = `E${index + 1}`;
      row[estimateTargetIdx] = issue.title; // セルにリンクは後で設定
      row[statusIdx] = `\
=LET(
  membersNames, ${membersTable.tableName}[${membersTable.headers.displayName}],
  membersResponseNecessities, ${membersTable.tableName}[${
        membersTable.headers.responseRequired
      }],
  membersEmails, ${membersTable.tableName}[${membersTable.headers.email}],
  estimatesEmails, ${formResponsesTable.tableName}[${
        formResponsesTable.headers.email
      }],
  estimatesPoints, ${formResponsesTable.tableName}[E${index + 1}. 見積り値],
  assigneeCount, COUNTIF(membersResponseNecessities, "<>不要"),
  respondedAssigneeCount, SUM(MAP(membersResponseNecessities, membersEmails, LAMBDA(n, e, IF(AND(n<>"不要", COUNTIF(estimatesEmails, e)), 1, 0)))),
  respondedMemberCount, SUM(MAP(membersEmails, LAMBDA(e, IF(COUNTIF(estimatesEmails, e), 1, 0)))),
  IF(
    assigneeCount > respondedAssigneeCount,
    "見積もり中",
    IF(
      AND(assigneeCount = 0, respondedMemberCount = 0),
      "必須回答者なし",
      LET(
        respondedMemberEstimates, FILTER(estimatesPoints, COUNTIF(membersEmails, estimatesEmails)),
        noSkipCount, COUNTIF(respondedMemberEstimates, "<>skip"),
        IF(
          noSkipCount = 0,
          "全員 skip",
          LET(
            points, FILTER(respondedMemberEstimates, respondedMemberEstimates<>"skip"),
            keys, {1;2;3;5;8;13;21;34;55;89},
            pos,  MAP(points, LAMBDA(v, IFERROR(MATCH(v, keys, 0), 0))),
            IF(
              SUM(--(pos=0))>0,
              "error",
              IF(MAX(pos)-MIN(pos) <= 2, "確定", "violation")
            )
          )
        )
      )
    )
  )
)
`;
      row[averageIdx] = `\
=LET(
  membersEmails, ${membersTable.tableName}[${membersTable.headers.email}],
  estimatesEmails, ${formResponsesTable.tableName}[${
        formResponsesTable.headers.email
      }],
  estimatesPoints, ${formResponsesTable.tableName}[E${index + 1}. 見積り値],
  respondedNoSkipMemberCount, SUM(MAP(estimatesEmails, estimatesPoints, LAMBDA(email, point, IF(AND(point<>"skip", COUNTIF(membersEmails, email)), 1, 0)))),
  IF(
    respondedNoSkipMemberCount = 0,
    "",
    ROUND(AVERAGE(FILTER(estimatesPoints, COUNTIF(membersEmails, estimatesEmails), estimatesPoints<>"skip")))
  )
)
`;
      row[responseSummaryIdx] = `\
=LET(
  membersEmails, ${membersTable.tableName}[${membersTable.headers.email}],
  membersNames,  ${membersTable.tableName}[${membersTable.headers.displayName}],
  estimatesEmails, ${formResponsesTable.tableName}[${
        formResponsesTable.headers.email
      }],
  estimatesPoints, ${formResponsesTable.tableName}[E${index + 1}. 見積り値],
  estimatesComments, ${formResponsesTable.tableName}[E${
        index + 1
      }. 見積もりの前提、質問],
  estimatesDisplayNames, MAP(estimatesEmails, LAMBDA(l, IFERROR(INDEX(membersNames, MATCH(l, membersEmails, 0)), l))),
  respondedMemberCount, SUM(MAP(estimatesEmails, LAMBDA(email, IF(COUNTIF(membersEmails, email), 1, 0)))),
  IF(
    respondedMemberCount = 0,
    "",
    LET(
      displayTexts, FILTER(
        MAP(
          estimatesDisplayNames,
          estimatesPoints,
          estimatesComments,
          LAMBDA(name, point, comment,
            SUBSTITUTE(
              SUBSTITUTE(
                SUBSTITUTE("（%name%）%point%: %comment%", "%name%", name),
                "%point%", IF(point = "skip", point, point & "P")
              ),
              "%comment%", comment
            )
          )
        ),
        estimatesEmails <> "dummy"
      ),
      TEXTJOIN(CHAR(10) & CHAR(10), TRUE, displayTexts)
    )
  )
)
`;
      row[minIdx] = `\
=LET(
  membersEmails, ${membersTable.tableName}[${membersTable.headers.email}],
  estimatesEmails, ${formResponsesTable.tableName}[${
        formResponsesTable.headers.email
      }],
  estimatesPoints, ${formResponsesTable.tableName}[E${index + 1}. 見積り値],
  respondedNoSkipMemberCount, SUM(MAP(estimatesEmails, estimatesPoints, LAMBDA(email, point, IF(AND(point<>"skip", COUNTIF(membersEmails, email)), 1, 0)))),
  IF(
    respondedNoSkipMemberCount = 0,
    "",
    MIN(FILTER(estimatesPoints, COUNTIF(membersEmails, estimatesEmails), estimatesPoints<>"skip"))
  )
)
`;
      row[maxIdx] = `\
=LET(
  membersEmails, ${membersTable.tableName}[${membersTable.headers.email}],
  estimatesEmails, ${formResponsesTable.tableName}[${
        formResponsesTable.headers.email
      }],
  estimatesPoints, ${formResponsesTable.tableName}[E${index + 1}. 見積り値],
  respondedNoSkipMemberCount, SUM(MAP(estimatesEmails, estimatesPoints, LAMBDA(email, point, IF(AND(point<>"skip", COUNTIF(membersEmails, email)), 1, 0)))),
  IF(
    respondedNoSkipMemberCount = 0,
    "",
    MAX(FILTER(estimatesPoints, COUNTIF(membersEmails, estimatesEmails), estimatesPoints<>"skip"))
  )
)
`;
      row[minByIdx] = `\
=LET(
  membersEmails, ${membersTable.tableName}[${membersTable.headers.email}],
  membersNames,  ${membersTable.tableName}[${membersTable.headers.displayName}],
  estimatesEmails, ${formResponsesTable.tableName}[${
        formResponsesTable.headers.email
      }],
  estimatesPoints, ${formResponsesTable.tableName}[E${index + 1}. 見積り値],
  respondedNoSkipMemberCount, SUM(MAP(estimatesEmails, estimatesPoints, LAMBDA(email, point, IF(AND(point<>"skip", COUNTIF(membersEmails, email)), 1, 0)))),
  IF(
    respondedNoSkipMemberCount = 0,
    "",
    LET(
      value, MIN(FILTER(estimatesPoints, COUNTIF(membersEmails, estimatesEmails), estimatesPoints<>"skip")),
      nameOrEmptyList, MAP(membersNames, membersEmails, LAMBDA(n, e, IF(COUNTIF(FILTER(estimatesEmails, estimatesPoints = value), e), n, ""))),
      names, FILTER(nameOrEmptyList, nameOrEmptyList <> ""),
      TEXTJOIN("、", TRUE, names)
    )
  )
)
`;
      row[maxByIdx] = `\
=LET(
  membersEmails, ${membersTable.tableName}[${membersTable.headers.email}],
  membersNames,  ${membersTable.tableName}[${membersTable.headers.displayName}],
  estimatesEmails, ${formResponsesTable.tableName}[${
        formResponsesTable.headers.email
      }],
  estimatesPoints, ${formResponsesTable.tableName}[E${index + 1}. 見積り値],
  respondedNoSkipMemberCount, SUM(MAP(estimatesEmails, estimatesPoints, LAMBDA(email, point, IF(AND(point<>"skip", COUNTIF(membersEmails, email)), 1, 0)))),
  IF(
    respondedNoSkipMemberCount = 0,
    "",
    LET(
      value, MAX(FILTER(estimatesPoints, COUNTIF(membersEmails, estimatesEmails), estimatesPoints<>"skip")),
      nameOrEmptyList, MAP(membersNames, membersEmails, LAMBDA(n, e, IF(COUNTIF(FILTER(estimatesEmails, estimatesPoints = value), e), n, ""))),
      names, FILTER(nameOrEmptyList, nameOrEmptyList <> ""),
      TEXTJOIN("、", TRUE, names)
    )
  )
)
`;

      return row;
    });

    // データ範囲のA1記法を作成
    const dataA1 = gridRangeToA1(
      {
        sheetId: resultSummaryMeta.sheetId,
        startRowIndex: dataStartRow,
        endRowIndex: dataStartRow + requiredRows,
        startColumnIndex: resultSummaryMeta.range.startColumnIndex,
        endColumnIndex: resultSummaryMeta.range.endColumnIndex,
      },
      resultSummaryMeta.sheetTitle
    );

    if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
      throw new Error("Sheets.Spreadsheets is not available");
    }
    Sheets.Spreadsheets.Values.update(
      { values: dataRows },
      spreadsheetId,
      dataA1,
      { valueInputOption: "USER_ENTERED" }
    );

    // 5. 見積もり対象列にリンクを設定
    const linkRequests = issueList.map((issue, index) => {
      const rowIndex = dataStartRow + index;
      if (!resultSummaryMeta.range) {
        throw new Error("ResultSummary table range is undefined");
      }
      const colIndex =
        (resultSummaryMeta.range.startColumnIndex || 0) + estimateTargetIdx;

      return {
        updateCells: {
          range: {
            sheetId: resultSummaryMeta.sheetId,
            startRowIndex: rowIndex,
            endRowIndex: rowIndex + 1,
            startColumnIndex: colIndex,
            endColumnIndex: colIndex + 1,
          },
          rows: [
            {
              values: [
                {
                  userEnteredValue: { stringValue: issue.title },
                  textFormatRuns: [
                    {
                      startIndex: 0,
                      format: { link: { uri: issue.url } },
                    },
                  ],
                },
              ],
            },
          ],
          fields: "userEnteredValue,textFormatRuns",
        },
      };
    });

    if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
      throw new Error("Sheets.Spreadsheets is not available");
    }
    Sheets.Spreadsheets.batchUpdate({ requests: linkRequests }, spreadsheetId);

    logInfo("Successfully updated ResultSummary table", {
      tableId: resultSummaryMeta.tableId,
      sheetId: resultSummaryMeta.sheetId,
      dataRowsInserted: requiredRows,
      newTableEndRow,
      dataA1,
      linksSet: linkRequests.length,
    });
  } catch (err) {
    const error = err instanceof Error ? err : new Error(String(err));
    logError("Failed to update ResultSummary table", {
      error: error.toString(),
      spreadsheetUrl,
    });
    throw error;
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
    requestSlack: "依頼 Slack メッセージ",
    completionSlack: "完了 Slack メッセージ",
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
    SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(url).build()
  );
};

/**
 * リッチテキストの要素
 * @typedef {{type: 'plain', text: string} | {type: 'link', text: string, url: string}} RichTextElement
 */

/** @typedef {{ elements: Array<RichTextElement> }} RichText */

/**
 * RichText を単純文字列に変換
 * @param {RichText} richText
 * @returns {string}
 */
const richTextToString = (richText) =>
  richText.elements.map((p) => p.text).join("");

/**
 * RichText を updateCells 用のセル値に変換
 * @param {RichText} richText
 */
const richTextToCell = (richText) => {
  const text = richTextToString(richText);
  /** @type {Array<{startIndex: number, format?: {link?: {uri: string}}}>} */
  const runs = [];
  let index = 0;
  for (const part of richText.elements) {
    if (part.type === "link" && part.text.length > 0) {
      runs.push({ startIndex: index, format: { link: { uri: part.url } } });
      const endIndex = index + part.text.length;
      if (endIndex < text.length) {
        runs.push({ startIndex: endIndex });
      }
    }
    index += part.text.length;
  }
  /** @type {{userEnteredValue: {stringValue: string}, textFormatRuns?: typeof runs}} */
  const cell = { userEnteredValue: { stringValue: text } };
  if (runs.length) {
    cell.textFormatRuns = runs;
  }
  return cell;
};

/**
 * 文字列全体がリンクの RichText セルを構築
 * @param {string} text
 * @param {string} url
 */
const buildLinkCell = (text, url) =>
  richTextToCell({ elements: [{ type: "link", text, url }] });

/**
 * テーブルの「データ先頭」（ヘッダー直下）に 1 行挿入し、値を書き込む。
 * - headerRowCount は 1 と仮定（現行UIの標準）
 * @param {{ date: string, midText: string, midUrl: string, formText: string, formUrl: string, resultText: string, resultUrl: string, requestSlackMessage: RichText, completionSlackMessage: RichText }} row
 */
const addEstimateHistoryTopRow = (row) => {
  const meta = getTableMetaByName(estimateHistoryTable.tableName);
  const {
    sheetId,
    sheetTitle,
    startCol0,
    endCol0,
    dataTop0,
    idxByName,
  } = getTableHeaderInfo(meta);
  const colCount = endCol0 - startCol0;

  const idxDate = idxByName(estimateHistoryTable.headers.date);
  const idxMid = idxByName(estimateHistoryTable.headers.mid);
  const idxForm = idxByName(estimateHistoryTable.headers.form);
  const idxResult = idxByName(estimateHistoryTable.headers.result);
  const idxRequestSlack = idxByName(estimateHistoryTable.headers.requestSlack);
  const idxCompletionSlack = idxByName(
    estimateHistoryTable.headers.completionSlack
  );

  /** @type {string[]} */
  const valuesRow = Array(colCount).fill("");
  valuesRow[idxDate] = row.date; // USER_ENTERED → 日付認識
  // ここでは表示文字列のみを書き込み、リンクは後段でセル自体に付与する
  valuesRow[idxMid] = row.midText;
  valuesRow[idxForm] = row.formText;
  valuesRow[idxResult] = row.resultText;
  valuesRow[idxRequestSlack] = richTextToString(row.requestSlackMessage);
  valuesRow[idxCompletionSlack] = richTextToString(row.completionSlackMessage);

  // 1) データ先頭に 1 行分のスペースを挿入（テーブル幅に限定）
  // 2) 直後にその行に値を書き込む
  const insertRange = {
    sheetId,
    startRowIndex: dataTop0,
    endRowIndex: dataTop0 + 1,
    startColumnIndex: startCol0,
    endColumnIndex: endCol0,
  };

  if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
    throw new Error("Sheets.Spreadsheets is not available");
  }
  Sheets.Spreadsheets.batchUpdate(
    {
      requests: [
        { insertRange: { range: insertRange, shiftDimension: "ROWS" } },
      ],
    },
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
  if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
    throw new Error("Sheets.Spreadsheets is not available");
  }
  Sheets.Spreadsheets.Values.update(
    { values: [valuesRow] },
    SpreadsheetApp.getActiveSpreadsheet().getId(),
    rowA1,
    { valueInputOption: "USER_ENTERED" }
  );

  // 3) セル自体にリンクを付与（HYPERLINK 式は使わない）
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
      {
        updateCells: {
          range: {
            sheetId,
            startRowIndex: dataTop0,
            endRowIndex: dataTop0 + 1,
            startColumnIndex: startCol0 + idxRequestSlack,
            endColumnIndex: startCol0 + idxRequestSlack + 1,
          },
          rows: [{ values: [richTextToCell(row.requestSlackMessage)] }],
          fields: "userEnteredValue,textFormatRuns",
        },
      },
      {
        updateCells: {
          range: {
            sheetId,
            startRowIndex: dataTop0,
            endRowIndex: dataTop0 + 1,
            startColumnIndex: startCol0 + idxCompletionSlack,
            endColumnIndex: startCol0 + idxCompletionSlack + 1,
          },
          rows: [{ values: [richTextToCell(row.completionSlackMessage)] }],
          fields: "userEnteredValue,textFormatRuns",
        },
      },
    ],
  };

  if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
    throw new Error("Sheets.Spreadsheets is not available");
  }
  Sheets.Spreadsheets.batchUpdate(
    linkReq,
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
    throw new Error("Sheets.Spreadsheets is not available");
  }
  const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
  const values = vr.values || [];
  if (!values.length || !values[0]) {
    throw new Error("テーブルが空です: 見積もり必要_課題リスト");
  }
  const { idxByName } = getTableHeaderInfo(meta, values);
  const titleIdx = idxByName(estimateIssueListTable.headers.title);
  const urlIdx = idxByName(estimateIssueListTable.headers.url);

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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
    throw new Error("Sheets.Spreadsheets is not available");
  }
  const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
  const values = vr.values || [];
  if (!values.length || !values[0]) {
    throw new Error("テーブルが空です: 見積もり必要_締切");
  }

  const { idxByName } = getTableHeaderInfo(meta, values);
  const dueIdx = idxByName(estimateDeadlineTable.headers.dueDate);

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
      const e = err instanceof Error ? err : new Error(String(err));
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


/** ===== Google Form セクション操作機能 =================== */

/**
 * Google FormのURLからFormオブジェクトを取得
 * @param {string} formUrl - Google FormのURL
 * @returns {any} - Formオブジェクト
 */
const getFormFromUrl = (formUrl) => {
  const formId = extractFormIdFromUrl(formUrl);
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
    if (items[i].getType() === FormApp.ItemType.PAGE_BREAK) {
      pageBreakIndices.push(i);
    }
  }

  if (pageBreakIndices.length === 0) {
    throw new Error("No PAGE_BREAK found in form");
  }

  const firstPageBreakIndex = pageBreakIndices[0];
  if (firstPageBreakIndex === undefined) {
    throw new Error("No PAGE_BREAK found");
  }
  const secondPageBreakIndex = pageBreakIndices[1];

  // 2つ目以降のPAGE_BREAKとそれ以降のアイテムを削除
  if (pageBreakIndices.length > 1) {
    logInfo("Removing extra PAGE_BREAK items", {
      extraPageBreaks: pageBreakIndices.length - 1,
      firstPageBreakIndex,
      secondPageBreakIndex,
    });

    if (secondPageBreakIndex !== undefined) {
      // 後ろから削除（インデックスがずれないように）
      for (let i = items.length - 1; i >= secondPageBreakIndex; i--) {
        form.deleteItem(i);
      }
    }

    // アイテムリストを再取得
    items = form.getItems();
  }

  const expectedStructure = [
    FormApp.ItemType.PAGE_BREAK,
    FormApp.ItemType.PARAGRAPH_TEXT,
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
    if (!issue) {
      throw new Error(`No issue found for section index ${i}`);
    }

    sectionHeaderLikeItem.setTitle(issue.url);
    premiseItem.setTitle(`E${i + 1}. 見積もりの前提、質問`);
    estimateItem.setTitle(`E${i + 1}. 見積り値`);

    logInfo(`Updated section ${i + 1} titles`, {
      sectionIndex: i + 1,
      issueTitle: issue.title,
      issueUrl: issue.url,
    });
  }

  // シートとの同期を待つ
  Utilities.sleep(2000);

  logInfo("Successfully updated all section titles with issue URLs", {
    totalSections: targetCount,
  });
};

/** ===== 追加: 見積もり必要_デバッグ ローダ =================== */
const estimateDebugTable = {
  tableName: "見積もり必要_デバッグ",
  headers: {
    key: "key",
    value: "value",
  },
};

/** @typedef {Record<string,string>} EstimateDebugMap */
/** @type {EstimateDebugMap|undefined} */
let _estimateDebugCache = undefined;

/**
 * 見積もり必要_デバッグ（テーブル）を読み込み、key-value マップを返す。
 * @returns {EstimateDebugMap}
 */
const getEstimateDebugMap = () => {
  if (_estimateDebugCache) {
    return _estimateDebugCache;
  }
  const meta = getTableMetaByName(estimateDebugTable.tableName);
  const a1 = gridRangeToA1(meta.range, meta.sheetTitle);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  if (!isSpreadsheetsCollection(Sheets.Spreadsheets)) {
    throw new Error(
      "Advanced Sheets API is not enabled. Please enable it in the Google Apps Script project."
    );
  }
  const vr = Sheets.Spreadsheets.Values.get(spreadsheetId, a1);
  const values = vr.values || [];
  if (!values.length || !values[0]) {
    throw new Error("テーブルが空です: 見積もり必要_デバッグ");
  }

  const { idxByName } = getTableHeaderInfo(meta, values);
  const keyIdx = idxByName(estimateDebugTable.headers.key);
  const valueIdx = idxByName(estimateDebugTable.headers.value);

  /** @type {EstimateDebugMap} */
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const k = String(row[keyIdx] ?? "").trim();
    const v = String(row[valueIdx] ?? "").trim();
    if (!k) {
      continue;
    }
    if (Object.prototype.hasOwnProperty.call(map, k) && map[k] !== v) {
      logWarn("duplicate key in 見積もり必要_デバッグ", {
        row: i + 1,
        key: k,
        prev: map[k],
        next: v,
      });
      continue;
    }
    map[k] = v;
  }

  _estimateDebugCache = map;
  logInfo("Loaded table 見積もり必要_デバッグ", {
    a1,
    countRows: values.length - 1,
    keys: Object.keys(map).length,
    tableId: meta.tableId,
  });
  return map;
};

/**
 * 指定された Google Form のタイトル、説明、および全 Item をログ出力する。
 * @param {string} formUrl - Google Form の URL
 * @returns {{title: string, description?: string, items: Array<{index:number,type:string,title:string,helpText?:string,choices?:string[]}>}}
 */
const debugLogForm = (formUrl) => {
  const form = getFormFromUrl(formUrl);
  const title = form.getTitle();
  const description = form.getDescription();
  logInfo("Form info", { title, description });
  const items = form.getItems();
  /** @type {Array<{index:number,type:string,title:string,helpText?:string,choices?:string[]}>} */
  const out = [];
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const type = item.getType();
    /** @type {{index:number,type:string,title:string,helpText?:string,choices?:string[]}} */
    const info = {
      index: i,
      type: String(type),
      title: item.getTitle(),
    };
    const helpText = item.getHelpText();
    if (helpText) {
      info.helpText = helpText;
    }
    if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
      info.choices = item
        .asMultipleChoiceItem()
        .getChoices()
        .map(
          /** @param {GoogleAppsScript.Forms.Choice} c */ (c) => c.getValue()
        );
    } else if (type === FormApp.ItemType.CHECKBOX) {
      info.choices = item
        .asCheckboxItem()
        .getChoices()
        .map(
          /** @param {GoogleAppsScript.Forms.Choice} c */ (c) => c.getValue()
        );
    } else if (type === FormApp.ItemType.LIST) {
      info.choices = item
        .asListItem()
        .getChoices()
        .map(
          /** @param {GoogleAppsScript.Forms.Choice} c */ (c) => c.getValue()
        );
    }
    out.push(info);
    logInfo("Form item", info);
  }
  return { title, description, items: out };
};

/** ===== エントリポイント（実行対象の公開） ============= */

/** 個別テスト実行 */
/**
 * @param {string} name
 */
const runTestByName = (name) =>
  safeMain("runTestByName", () => runTestsCore({ names: [name] }));

/** 指定名の複数テストを実行 */
/**
 * @param {string[]} names
 */
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

/**
 * スプレッドシートが開かれたときに実行される関数
 * カスタムメニューを追加する
 */
const onOpen = () => {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("拡張コマンド")
    .addItem("新規 async 見積もり発行", "runCreateEstimate")
    .addToUi();
};

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
    const issueList = getEstimateIssueList().filter(({ title, url }) => title || url);
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

/**
 * デバッグ用: テーブル「見積もり必要_デバッグ」を読み込み logInfo する。
 * 使用例: runDebugEstimateDebugTable()
 */
const runDebugEstimateDebugTable = () =>
  safeMain("runDebugEstimateDebugTable", () => {
    const map = getEstimateDebugMap();
    logInfo("Estimate debug map", map);
    return map;
  });

/**
 * デバッグ用: 見積もり必要_デバッグテーブルの google-form-url の Form をログ出力。
 * 使用例: runDebugForm()
 */
const runDebugForm = () =>
  safeMain("runDebugForm", () => {
    const map = getEstimateDebugMap();
    const formUrl = map["google-form-url"];
    if (!formUrl) {
      throw new Error(
        'key "google-form-url" not found in 見積もり必要_デバッグ'
      );
    }
    logInfo("Logging form", { formUrl });
    return debugLogForm(formUrl);
  });

