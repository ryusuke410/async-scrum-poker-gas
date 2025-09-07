import type { sheets_v4 } from "googleapis";

// 使う分だけ alias（必要に応じて増やしてください）
type VR = sheets_v4.Schema$ValueRange;
type UpdateValuesResponse = sheets_v4.Schema$UpdateValuesResponse;
type BatchUpdateSpreadsheetRequest =
  sheets_v4.Schema$BatchUpdateSpreadsheetRequest;
type BatchUpdateSpreadsheetResponse =
  sheets_v4.Schema$BatchUpdateSpreadsheetResponse;
type P_Get = sheets_v4.Params$Resource$Spreadsheets$Values$Get;
type P_Update = sheets_v4.Params$Resource$Spreadsheets$Values$Update;

// declare var Sheets: GoogleAppsScript.Sheets;

declare global {
  namespace GoogleAppsScript {
    namespace Sheets {
      interface SpreadsheetsValues {
        get(
          p: sheets_v4.Params$Resource$Spreadsheets$Values$Get,
        ): sheets_v4.Schema$ValueRange;
        update(
          p: sheets_v4.Params$Resource$Spreadsheets$Values$Update,
        ): sheets_v4.Schema$UpdateValuesResponse;
      }
      interface Spreadsheets {
        Values: SpreadsheetsValues;
        batchUpdate(p: {
          spreadsheetId: string;
          resource: sheets_v4.Schema$BatchUpdateSpreadsheetRequest;
        }): sheets_v4.Schema$BatchUpdateSpreadsheetResponse;
      }
      namespace Schema {
        interface Sheet {
          tables?: Sheets.Schema.Table[] | undefined;
        }
        interface Table {
          tableId?: string | undefined;
          name?: string | undefined;
          range?: Sheets.Schema.GridRange | undefined;
          rowsProperties?: Sheets.Schema.TableRowProperties | undefined;
          columnsProperties?: Sheets.Schema.TableColumnProperties | undefined;
        }
        interface TableRowProperties {
          headerColorStyle?: Sheets.Schema.ColorStyle | undefined;
          firstBandColorStyle?: Sheets.Schema.ColorStyle | undefined;
          secondBandColorStyle?: Sheets.Schema.ColorStyle | undefined;
          footerColorStyle?: Sheets.Schema.ColorStyle | undefined;
        }
        interface TableColumnProperties {
          columnIndex?: number | undefined;
          columnName?: string | undefined;
          columnType?: Sheets.Schema.ColumnType;
          dataValidationRule?: Sheets.Schema.DataValidationRule | undefined;
        }
        type ColumnType =
          | "COLUMN_TYPE_UNSPECIFIED"
          | "DOUBLE"
          | "CURRENCY"
          | "PERCENT"
          | "DATE"
          | "TIME"
          | "DATE_TIME"
          | "TEXT"
          | "BOOLEAN"
          | "DROPDOWN"
          | "FILES_CHIP"
          | "PEOPLE_CHIP"
          | "FINANCE_CHIP"
          | "PLACE_CHIP"
          | "RATINGS_CHIP";
        interface Request {
          updateTable?: Sheets.Schema.UpdateTableRequest | undefined;
        }
        interface UpdateTableRequest {
          table?: Sheets.Schema.Table | undefined;
          fields?: string | undefined;
        }
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
