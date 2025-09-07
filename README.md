# Async Scrum Poker (GAS)

A Google Apps Script that automates asynchronous Scrum Poker using Google Sheets and Google Forms. It copies template assets, links a Form to a Spreadsheet, prepares sections per issue, and shares artifacts with the right members.

## Quick Start

- Open the target Google Spreadsheet and launch Extensions → Apps Script.
- Create a script file and paste the contents of `code.gas.js`.
- Enable Advanced Services: in Apps Script, turn on “Google Sheets API” (and ensure it’s enabled in the linked Google Cloud project). The script also uses `DriveApp`, `SpreadsheetApp`, and `FormApp`.
- Save the project and refresh the spreadsheet. A custom menu “拡張コマンド” will appear.
- Run: 拡張コマンド → “新規 async 見積もり発行” to generate a new estimation set.

## Spreadsheet Setup

- Provide a table named "見積もり必要_テンプレート" with headers "名前" and "リンク". Include rows:
  - 名前: `Google Form` → リンク: template Form URL
  - 名前: `中間スプシ` → リンク: intermediate Spreadsheet URL
  - 名前: `結果スプシ` → リンク: result Spreadsheet URL
- The script reads this table to copy templates and wire everything together.

## Local Development (optional)

- Toolchain: `mise install` (Node 22, pnpm 10), then `pnpm i`.
- Type check: `pnpm exec tsc --noEmit` (project uses `// @ts-check` and strict options in `tsconfig.json`).
- Style and design principles live in `coding-guide.md`.

## Testing

- Tests execute inside Apps Script. From the editor, run any of the exposed functions at the bottom of `code.gas.js`, for example:
  - `testCore` (side-effect free core checks)
  - `testTemplateCore`, `testPoMembersCore`, or individual helpers like `testSampleTrue`
- Results are logged via `Logger.log` and summarized after each run.

## Notes

- Do not paste secrets. Share Forms/Sheets with least privilege.
- First runs may request authorization for Sheets, Drive, and Forms scopes.
