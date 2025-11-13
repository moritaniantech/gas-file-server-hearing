/**
 * @fileoverview
 * (ファイル2/2) MIXI (新)project-2 シート処理用スクリプト
 * * このファイルには「project用」の処理のみが含まれます。
 * * 共通メニュー(onOpen)は Code_div.gs に記載されています。
 */

// ===================================================
// project フォルダ用 設定
// ===================================================

/** (project) 元データが記載されているシート名 (ファイルリストから '（新）project-2' と推定) */
const PROJECT_SOURCE_SHEET_NAME = '（新）project-2';

/** (project) 作成したシートのURL一覧を記載するシート名 */
const PROJECT_SUMMARY_SHEET_NAME = 'project回答シートURL一覧';

/** (project) ファイル数が0で除外されたデータをまとめるスプレッドシートのファイル名 */
const PROJECT_EXCLUDED_SHEET_NAME = 'project確認先を除外';

/** (project) L列・M列が両方 '×' だった場合のファイル名 */
const PROJECT_UNKNOWN_SHEET_NAME = 'project不明';

/** (project)【重要】回答用スプレッドシートを保存するGoogle DriveフォルダのID */
const PROJECT_DESTINATION_FOLDER_ID = '1OjgFtdJYA3kyZm95ogtiiYAkU7zLL8wM'; // ← ★★★ 設定してください ★★★

// (project) (新)project-2シートの列インデックス (divと同じと想定)
const PROJECT_COL_A_DIV1 = 0;
const PROJECT_COL_B_DIV2 = 1;
const PROJECT_COL_D_FOLDER_COUNT = 3;
const PROJECT_COL_E_FILE_COUNT = 4;
const PROJECT_COL_F_DATA_SIZE = 5;
const PROJECT_COL_G_LAST_UPDATED = 6;
const PROJECT_COL_H_HONBU = 7;
const PROJECT_COL_I_BUSHITSU = 8;
const PROJECT_COL_J_MIGRATION_DEST = 9;
const PROJECT_COL_K_MIGRATION_METHOD = 10;
const PROJECT_COL_L_HONBUCHO = 11;
const PROJECT_COL_M_BUSHITSUCHO = 12;


// ===================================================
// project フォルダ用 スクリプト
// ===================================================

/**
 * (project) メイン処理1: (新)project-2シートを読み込み、確認先ごとにスプレッドシートを分割作成します。
 */
function project_createResponseSheets() {
  if (PROJECT_DESTINATION_FOLDER_ID === 'YOUR_FOLDER_ID_HERE_FOR_PROJECT') {
    SpreadsheetApp.getUi().alert('スクリプトエラー (project)', 'PROJECT_DESTINATION_FOLDER_IDが設定されていません。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(PROJECT_SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(`エラー (project): シート「${PROJECT_SOURCE_SHEET_NAME}」が見つかりません。`);
    return;
  }

  let destinationFolder;
  try {
    destinationFolder = DriveApp.getFolderById(PROJECT_DESTINATION_FOLDER_ID);
  } catch (e) {
    SpreadsheetApp.getUi().alert(`エラー (project): 指定されたフォルダID「${PROJECT_DESTINATION_FOLDER_ID}」が見つかりません。${e.message}`);
    return;
  }

  const data = sourceSheet.getDataRange().getValues();
  const headers = data.shift(); // ヘッダー行を取得

  const dataByPerson = {};
  const excludedData = [];
  
  Logger.log(`(project) 元データ ${data.length}行の処理を開始します。`);

  // --- (project) データの振り分け ---
  data.forEach((row, index) => {
    if (row.join('').length === 0) return;
    const fileCount = row[PROJECT_COL_E_FILE_COUNT];

    if (fileCount === 0 || fileCount === '0') {
      excludedData.push(row);
      return;
    }

    const personL = row[PROJECT_COL_L_HONBUCHO];
    const personM = row[PROJECT_COL_M_BUSHITSUCHO];
    let confirmationPerson = PROJECT_UNKNOWN_SHEET_NAME;

    if (personL && personL !== '×' && personM === '×') {
      confirmationPerson = personL;
    } else if (personL === '×' && personM && personM !== '×') {
      confirmationPerson = personM;
    } else if (personL && personL !== '×' && personM && personM !== '×') {
      confirmationPerson = personM;
    } else if (personL === '×' && personM === '×') {
      confirmationPerson = PROJECT_UNKNOWN_SHEET_NAME;
    }

    // (project) v8修正: G/H列 と J/K列 の「値」を入れ替え
    const newRow = [
      row[PROJECT_COL_A_DIV1], // 0: A
      row[PROJECT_COL_B_DIV2], // 1: B
      row[PROJECT_COL_D_FOLDER_COUNT], // 2: C
      row[PROJECT_COL_E_FILE_COUNT], // 3: D
      row[PROJECT_COL_F_DATA_SIZE], // 4: E
      row[PROJECT_COL_G_LAST_UPDATED], // 5: F
      row[PROJECT_COL_J_MIGRATION_DEST], // 6: G (★元データJ)
      row[PROJECT_COL_K_MIGRATION_METHOD], // 7: H (★元データK)
      '', // 8: I
      row[PROJECT_COL_H_HONBU], // 9: J (★元データH)
      row[PROJECT_COL_I_BUSHITSU], // 10: K (★元データI)
      '', // 11: L
      false, // 12: M
      false, // 13: N
      '' // 14: O
    ];

    if (!dataByPerson[confirmationPerson]) {
      dataByPerson[confirmationPerson] = [];
    }
    dataByPerson[confirmationPerson].push(newRow);
  });
  
  Logger.log(`(project) データ振り分け完了。確認先: ${Object.keys(dataByPerson).length}件、除外: ${excludedData.length}件`);

  // --- (project) 回答用シートのヘッダー定義 (v8) ---
  const outputHeaders = [
    'projectフォルダ_1階層目', 'projectフォルダ 2階層目', 'フォルダ数', 'ファイル数', 'データ容量/GB',
    '最終更新日', '対象本部 ※推測込', '対象部室 ※ 推測込', '回答者メールアドレス', '移行先',
    '移行方法', '共有ドライブ名', '個人情報有無', '自動化有無 ※スクリプトやRPAなど', 'その他'
  ];
  // ユーザーが入力する列 (v8)
  const userInputHeaderIndices = [8, 9, 10, 11, 12, 13, 14];

  // --- (project) データバリデーションルール ---
  const rules = {
    migrationRule: SpreadsheetApp.newDataValidation().requireValueInList(['Google Drive', 'AWS S3', '別テナント', '不要'], true).build(),
    methodRule: SpreadsheetApp.newDataValidation().requireValueInList(['自対応', 'hatakan依頼', '不要'], true).build(),
    checkboxRule: SpreadsheetApp.newDataValidation().requireCheckbox().build()
  };

  // --- (project) URL一覧シートの準備 ---
  let summarySheet = ss.getSheetByName(PROJECT_SUMMARY_SHEET_NAME);
  if (summarySheet) {
    summarySheet.clear();
  } else {
    summarySheet = ss.insertSheet(PROJECT_SUMMARY_SHEET_NAME);
  }
  const summaryData = [['確認先氏名', 'SpreadsheetのURL', '件数']];

  // --- (project) 確認先ごとにスプレッドシートを作成 ---
  for (const person in dataByPerson) {
    const fileName = `(project) ${person}`; // ファイル名に(project)を追加
    const rows = dataByPerson[person];
    Logger.log(`(project) シート作成開始: ${fileName} (${rows.length}件)`);
    try {
      // 内部ヘルパー関数を呼び出す
      const result = project_createAndFormatSheet(fileName, outputHeaders, rows, destinationFolder, userInputHeaderIndices, rules);
      summaryData.push([person, result.url, result.rowCount]);
      Logger.log(`(project) 作成完了: ${person} (URL: ${result.url})`);
    } catch (e) {
      Logger.log(`(project) エラー: ${person} のシート作成に失敗しました。 ${e.message}`);
      summaryData.push([person, `作成失敗: ${e.message}`, 0]);
    }
  }

  // --- (project) 「除外」シートの作成 ---
  if (excludedData.length > 0) {
    const fileName = PROJECT_EXCLUDED_SHEET_NAME;
    Logger.log(`(project) シート作成開始: ${fileName} (${excludedData.length}件)`);
    const mappedExcludedData = excludedData.map(row => [
      row[PROJECT_COL_A_DIV1], row[PROJECT_COL_B_DIV2], row[PROJECT_COL_D_FOLDER_COUNT],
      row[PROJECT_COL_E_FILE_COUNT], row[PROJECT_COL_F_DATA_SIZE], row[PROJECT_COL_G_LAST_UPDATED],
      row[PROJECT_COL_J_MIGRATION_DEST], row[PROJECT_COL_K_MIGRATION_METHOD], '',
      row[PROJECT_COL_H_HONBU], row[PROJECT_COL_I_BUSHITSU],
      '', false, false, ''
    ]);
    
    try {
      const result = project_createAndFormatSheet(fileName, outputHeaders, mappedExcludedData, destinationFolder, userInputHeaderIndices, rules);
      summaryData.push([fileName, result.url, result.rowCount]);
      Logger.log(`(project) 作成完了: ${fileName} (URL: ${result.url})`);
    } catch (e) {
      Logger.log(`(project) エラー: ${fileName} のシート作成に失敗しました。 ${e.message}`);
      summaryData.push([fileName, `作成失敗: ${e.message}`, 0]);
    }
  }

  summarySheet.getRange(1, 1, summaryData.length, 3).setValues(summaryData);
  summarySheet.autoResizeColumns(1, 3);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('(project) 回答シート作成処理が完了しました。');
}

/**
 * (project) ヘルパー関数: スプレッドシートを作成し、フォーマットします。
 */
function project_createAndFormatSheet(fileName, headers, dataRows, folder, highlightIndices, rules) {
  const newSs = SpreadsheetApp.create(fileName);
  const file = DriveApp.getFileById(newSs.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  const sheet = newSs.getSheets()[0];
  sheet.setName('回答');
  
  const numRows = dataRows.length;
  const numCols = headers.length;

  if (sheet.getMaxRows() > numRows + 1) {
    sheet.deleteRows(numRows + 2, sheet.getMaxRows() - (numRows + 1));
  }
  if (sheet.getMaxColumns() > numCols) {
    sheet.deleteColumns(numCols + 1, sheet.getMaxColumns() - numCols);
  }

  // (グリッド線非表示はエラーのため削除)

  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setValues([headers]);
  headerRange.setBackground('#d9d9d9').setFontWeight('bold').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  highlightIndices.forEach(colIndex => {
    sheet.getRange(1, colIndex + 1).setBackground('#c9daf8');
  });

  if (numRows > 0) {
    Logger.log(`> (project) ${fileName}: ${numRows}件のデータを書き込みます。`);
    try {
      const dataRange = sheet.getRange(2, 1, numRows, numCols);
      dataRange.setValues(dataRows);
      dataRange.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID); // 細線

      // (project) データバリデーション設定 (v8)
      sheet.getRange(2, 10, numRows, 1).setDataValidation(rules.migrationRule); // J列
      sheet.getRange(2, 11, numRows, 1).setDataValidation(rules.methodRule); // K列
      sheet.getRange(2, 13, numRows, 1).setDataValidation(rules.checkboxRule); // M列
      sheet.getRange(2, 14, numRows, 1).setDataValidation(rules.checkboxRule); // N列

    } catch (e) {
      Logger.log(`> (project) ${fileName}: データ書き込みまたはフォーマット中にエラー: ${e.message}`);
    }
  } else {
    Logger.log(`> (project) ${fileName}: データ件数が0のため、スキップしました。`);
  }

  try {
    sheet.autoResizeColumns(1, numCols);
  } catch (e) {
    Logger.log(`> (project) ${fileName}: 列幅の自動調整に失敗しました。 ${e.message}`);
  }

  return { url: newSs.getUrl(), rowCount: numRows };
}

/**
 * (project) メイン処理2: URL一覧シートを読み込み、各シートの回答をマージします。
 */
function project_mergeResponseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName(PROJECT_SUMMARY_SHEET_NAME);
  if (!summarySheet) {
    SpreadsheetApp.getUi().alert(`エラー (project): 「${PROJECT_SUMMARY_SHEET_NAME}」が見つかりません。`);
    return;
  }

  const urlData = summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 3).getValues();
  const mergedData = [];
  let headers = [];
  let headersSet = false;

  for (const row of urlData) {
    const personName = row[0];
    const url = row[1];
    if (!url || !url.startsWith('http')) {
      Logger.log(`(project) スキップ: ${personName} (無効なURL)`);
      continue;
    }

    try {
      const targetSs = SpreadsheetApp.openByUrl(url);
      const targetSheet = targetSs.getSheets()[0];
      const values = targetSheet.getDataRange().getValues();

      if (!headersSet) {
        headers = values.shift();
        headers.unshift('確認先');
        mergedData.push(headers);
        headersSet = true;
      } else {
        values.shift();
      }

      values.forEach(dataRow => {
        dataRow.unshift(personName);
        mergedData.push(dataRow);
      });
      Logger.log(`(project) マージ完了: ${personName}`);

    } catch (e) {
      Logger.log(`(project) エラー: ${personName} のシート読み込み失敗。 (URL: ${url}) ${e.message}`);
      if (headersSet) {
        const errorRow = new Array(headers.length).fill('');
        errorRow[0] = personName;
        errorRow[1] = `シートの読み込みに失敗しました: ${e.message}`;
        mergedData.push(errorRow);
      }
    }
  }

  if (mergedData.length <= 1) { // 1 = ヘッダーのみ
    Logger.log('(project) マージするデータがありませんでした。');
    SpreadsheetApp.getUi().alert('(project) マージするデータがありませんでした。');
    return;
  }

  const outputSheetName = 'projectマージ済み回答';
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = ss.insertSheet(outputSheetName);
  }

  outputSheet.getRange(1, 1, mergedData.length, mergedData[0].length).setValues(mergedData);
  outputSheet.autoResizeColumns(1, mergedData[0].length);
  outputSheet.setFrozenRows(1);
  outputSheet.getRange(1, 1, 1, mergedData[0].length).setFontWeight('bold');
  Logger.log('(project) マージ処理が完了しました。');
  SpreadsheetApp.getUi().alert('(project) 回答シートマージ処理が完了しました。');
}