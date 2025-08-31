import type { sheets_v4 } from "googleapis";

// 使う分だけ alias（必要に応じて増やしてください）
type VR = sheets_v4.Schema$ValueRange;
type UpdateValuesResponse = sheets_v4.Schema$UpdateValuesResponse;
type BatchUpdateSpreadsheetRequest = sheets_v4.Schema$BatchUpdateSpreadsheetRequest;
type BatchUpdateSpreadsheetResponse = sheets_v4.Schema$BatchUpdateSpreadsheetResponse;
type P_Get = sheets_v4.Params$Resource$Spreadsheets$Values$Get;
type P_Update = sheets_v4.Params$Resource$Spreadsheets$Values$Update;

// declare var Sheets: GoogleAppsScript.Sheets;

declare global {
  namespace GoogleAppsScript {
    namespace Sheets {
      interface SpreadsheetsValues {
        get(p: sheets_v4.Params$Resource$Spreadsheets$Values$Get): sheets_v4.Schema$ValueRange;
        update(p: sheets_v4.Params$Resource$Spreadsheets$Values$Update): sheets_v4.Schema$UpdateValuesResponse;
      }
      interface Spreadsheets {
        Values: SpreadsheetsValues;
        batchUpdate(p: {
          spreadsheetId: string;
          resource: sheets_v4.Schema$BatchUpdateSpreadsheetRequest;
        }): sheets_v4.Schema$BatchUpdateSpreadsheetResponse;
      }
    }
  }
}

// // GAS の Advanced Service 実体に“型”を当てる（実行時は既存の Sheets を使う）
// declare const Sheets: {
//   Spreadsheets: {
//     Values: {
//       get(params: P_Get): VR;
//       update(params: P_Update): UpdateValuesResponse;
//     };
//     batchUpdate(p: {
//       spreadsheetId: string;
//       resource: BatchUpdateSpreadsheetRequest;
//     }): BatchUpdateSpreadsheetResponse;
//   };
// };

export {};
