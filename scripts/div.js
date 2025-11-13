/**
 * @fileoverview
 * (ファイル1/2) MIXI (新)div シート処理用スクリプト
 * * このファイルには「div用」の処理と、「共通メニュー(onOpen)」が含まれます。
 */

// ===================================================
// 共通メニュー (divとprojectを両方呼び出す)
// ===================================================

/**
 * Googleスプレッドシートを開いたときにカスタムメニューを追加します。
 * (このonOpen関数はプロジェクト全体で1つだけにしてください)
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GAS実行メニュー')
    .addSubMenu(ui.createMenu('div フォルダ処理')
      .addItem('1. (div) 回答シート作成', 'div_createResponseSheets')
      .addItem('2. (div) 回答シートマージ', 'div_mergeResponseSheets'))
    .addSeparator() // 区切り線
    .addSubMenu(ui.createMenu('project フォルダ処理')
      .addItem('1. (project) 回答シート作成', 'project_createResponseSheets')
      .addItem('2. (project) 回答シートマージ', 'project_mergeResponseSheets'))
    .addToUi();
}


// ===================================================
// div フォルダ用 設定
// ===================================================

/** (div) 元データが記載されているシート名 */
const DIV_SOURCE_SHEET_NAME = '（新）div';

/** (div) 作成したシートのURL一覧を記載するシート名 */
const DIV_SUMMARY_SHEET_NAME = 'div回答シートURL一覧';

/** (div) ファイル数が0で除外されたデータをまとめるスプレッドシートのファイル名 */
const DIV_EXCLUDED_SHEET_NAME = 'div確認先を除外';

/** (div) L列・M列が両方 '×' だった場合のファイル名 */
const DIV_UNKNOWN_SHEET_NAME = 'div不明';

/** (div)【重要】回答用スプレッドシートを保存するGoogle DriveフォルダのID */
const DIV_DESTINATION_FOLDER_ID = '1gCL4fPnyGpdaWP3eHFVdGvZGMWWXVI1j'; // ← ★★★ 設定してください ★★★

// (div) (新)divシートの列インデックス
const DIV_COL_A_DIV1 = 0;
const DIV_COL_B_DIV2 = 1;
const DIV_COL_D_FOLDER_COUNT = 3;
const DIV_COL_E_FILE_COUNT = 4;
const DIV_COL_F_DATA_SIZE = 5;
const DIV_COL_G_LAST_UPDATED = 6;
const DIV_COL_H_HONBU = 7;
const DIV_COL_I_BUSHITSU = 8;
const DIV_COL_J_MIGRATION_DEST = 9;
const DIV_COL_K_MIGRATION_METHOD = 10;
const DIV_COL_L_HONBUCHO = 11;
const DIV_COL_M_BUSHITSUCHO = 12;


// ===================================================
// div フォルダ用 スクリプト
// ===================================================

/**
 * (div) メイン処理1: (新)divシートを読み込み、確認先ごとにスプレッドシートを分割作成します。
 */
function div_createResponseSheets() {
  // ユーザー認証チェック
  const allowedUser = 'ryosuke.morita.ts@mixi.co.jp';
  const currentUser = Session.getActiveUser().getEmail();
  if (currentUser !== allowedUser) {
    SpreadsheetApp.getUi().alert('エラー (div)', `このスクリプトは ${allowedUser} のみが実行できます。\n現在のユーザー: ${currentUser}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  if (DIV_DESTINATION_FOLDER_ID === 'YOUR_FOLDER_ID_HERE_FOR_DIV') {
    SpreadsheetApp.getUi().alert('スクリプトエラー (div)', 'DIV_DESTINATION_FOLDER_IDが設定されていません。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(DIV_SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(`エラー (div): シート「${DIV_SOURCE_SHEET_NAME}」が見つかりません。`);
    return;
  }

  let destinationFolder;
  try {
    destinationFolder = DriveApp.getFolderById(DIV_DESTINATION_FOLDER_ID);
  } catch (e) {
    SpreadsheetApp.getUi().alert(`エラー (div): 指定されたフォルダID「${DIV_DESTINATION_FOLDER_ID}」が見つかりません。${e.message}`);
    return;
  }

  const data = sourceSheet.getDataRange().getValues();
  const headers = data.shift(); // ヘッダー行を取得

  const dataByPerson = {};
  const excludedData = [];
  
  Logger.log(`(div) 元データ ${data.length}行の処理を開始します。`);

  // --- (div) データの振り分け ---
  data.forEach((row, index) => {
    if (row.join('').length === 0) return;
    const fileCount = row[DIV_COL_E_FILE_COUNT];

    if (fileCount === 0 || fileCount === '0') {
      excludedData.push(row);
      return;
    }

    // AWS S3が選択されている場合は除外シートに追加
    const migrationDest = row[DIV_COL_J_MIGRATION_DEST];
    if (migrationDest === 'AWS S3') {
      excludedData.push(row);
      return;
    }

    const personL = row[DIV_COL_L_HONBUCHO];
    const personM = row[DIV_COL_M_BUSHITSUCHO];
    let confirmationPerson = DIV_UNKNOWN_SHEET_NAME;

    if (personL && personL !== '×' && personM === '×') {
      confirmationPerson = personL;
    } else if (personL === '×' && personM && personM !== '×') {
      confirmationPerson = personM;
    } else if (personL && personL !== '×' && personM && personM !== '×') {
      confirmationPerson = personM;
    } else if (personL === '×' && personM === '×') {
      confirmationPerson = DIV_UNKNOWN_SHEET_NAME;
    }

    // (div) v8修正: G/H列 と J/K列 の「値」を入れ替え
    const newRow = [
      row[DIV_COL_A_DIV1], // 0: A
      row[DIV_COL_B_DIV2], // 1: B
      row[DIV_COL_D_FOLDER_COUNT], // 2: C
      row[DIV_COL_E_FILE_COUNT], // 3: D
      row[DIV_COL_F_DATA_SIZE], // 4: E
      row[DIV_COL_G_LAST_UPDATED], // 5: F
      row[DIV_COL_J_MIGRATION_DEST], // 6: G (★元データJ)
      row[DIV_COL_K_MIGRATION_METHOD], // 7: H (★元データK)
      '', // 8: I
      row[DIV_COL_H_HONBU], // 9: J (★元データH)
      row[DIV_COL_I_BUSHITSU], // 10: K (★元データI)
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
  
  Logger.log(`(div) データ振り分け完了。確認先: ${Object.keys(dataByPerson).length}件、除外: ${excludedData.length}件`);

  // --- (div) 回答用シートのヘッダー定義 (v8) ---
  const outputHeaders = [
    'divフォルダ_1階層目', 'divフォルダ 2階層目', 'フォルダ数', 'ファイル数', 'データ容量/GB',
    '最終更新日', '対象本部 ※推測込', '対象部室 ※ 推測込', '回答者メールアドレス', '移行先',
    '移行方法', '共有ドライブ名', '個人情報有無', '自動化有無 ※スクリプトやRPAなど', 'その他'
  ];
  const userInputHeaderIndices = [8, 9, 10, 11, 12, 13, 14];

  // --- (div) データバリデーションルール ---
  const rules = {
    migrationRule: SpreadsheetApp.newDataValidation().requireValueInList(['Googleドライブ', 'AWS S3', '別テナント', '不要'], true).build(),
    methodRule: SpreadsheetApp.newDataValidation().requireValueInList(['自対応', 'hatakan依頼', '不要'], true).build(),
    checkboxRule: SpreadsheetApp.newDataValidation().requireCheckbox().build()
  };

  // --- (div) URL一覧シートの準備 ---
  let summarySheet = ss.getSheetByName(DIV_SUMMARY_SHEET_NAME);
  if (summarySheet) {
    summarySheet.clear();
  } else {
    summarySheet = ss.insertSheet(DIV_SUMMARY_SHEET_NAME);
  }
  const summaryData = [['確認先氏名', 'SpreadsheetのURL', '件数']];

  // --- (div) 確認先ごとにスプレッドシートを作成 ---
  for (const person in dataByPerson) {
    const fileName = `(div) ${person}`; // ファイル名に(div)を追加
    const rows = dataByPerson[person];
    Logger.log(`(div) シート作成開始: ${fileName} (${rows.length}件)`);
    try {
      // 内部ヘルパー関数を呼び出す
      const result = div_createAndFormatSheet(fileName, outputHeaders, rows, destinationFolder, userInputHeaderIndices, rules);
      summaryData.push([person, result.url, result.rowCount]);
      Logger.log(`(div) 作成完了: ${person} (URL: ${result.url})`);
    } catch (e) {
      Logger.log(`(div) エラー: ${person} のシート作成に失敗しました。 ${e.message}`);
      summaryData.push([person, `作成失敗: ${e.message}`, 0]);
    }
  }

  // --- (div) 「除外」シートの作成 ---
  if (excludedData.length > 0) {
    const fileName = DIV_EXCLUDED_SHEET_NAME;
    Logger.log(`(div) シート作成開始: ${fileName} (${excludedData.length}件)`);
    const mappedExcludedData = excludedData.map(row => [
      row[DIV_COL_A_DIV1], row[DIV_COL_B_DIV2], row[DIV_COL_D_FOLDER_COUNT],
      row[DIV_COL_E_FILE_COUNT], row[DIV_COL_F_DATA_SIZE], row[DIV_COL_G_LAST_UPDATED],
      row[DIV_COL_J_MIGRATION_DEST], row[DIV_COL_K_MIGRATION_METHOD], '',
      row[DIV_COL_H_HONBU], row[DIV_COL_I_BUSHITSU],
      '', false, false, ''
    ]);
    
    try {
      const result = div_createAndFormatSheet(fileName, outputHeaders, mappedExcludedData, destinationFolder, userInputHeaderIndices, rules);
      summaryData.push([fileName, result.url, result.rowCount]);
      Logger.log(`(div) 作成完了: ${fileName} (URL: ${result.url})`);
    } catch (e) {
      Logger.log(`(div) エラー: ${fileName} のシート作成に失敗しました。 ${e.message}`);
      summaryData.push([fileName, `作成失敗: ${e.message}`, 0]);
    }
  }

  summarySheet.getRange(1, 1, summaryData.length, 3).setValues(summaryData);
  summarySheet.autoResizeColumns(1, 3);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('(div) 回答シート作成処理が完了しました。');
}

/**
 * (div) ヘルパー関数: スプレッドシートを作成し、フォーマットします。
 */
function div_createAndFormatSheet(fileName, headers, dataRows, folder, highlightIndices, rules) {
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
    Logger.log(`> (div) ${fileName}: ${numRows}件のデータを書き込みます。`);
    try {
      const dataRange = sheet.getRange(2, 1, numRows, numCols);
      dataRange.setValues(dataRows);
      dataRange.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID); // 細線

      // (div) データバリデーション設定 (v8)
      sheet.getRange(2, 10, numRows, 1).setDataValidation(rules.migrationRule); // J列
      sheet.getRange(2, 11, numRows, 1).setDataValidation(rules.methodRule); // K列
      sheet.getRange(2, 13, numRows, 1).setDataValidation(rules.checkboxRule); // M列
      sheet.getRange(2, 14, numRows, 1).setDataValidation(rules.checkboxRule); // N列

    } catch (e) {
      Logger.log(`> (div) ${fileName}: データ書き込みまたはフォーマット中にエラー: ${e.message}`);
    }
  } else {
    Logger.log(`> (div) ${fileName}: データ件数が0のため、スキップしました。`);
  }

  // 列幅の調整（列名と値の両方を考慮）
  try {
    sheet.autoResizeColumns(1, numCols);
    SpreadsheetApp.flush(); // 自動調整を確実に反映
    
    // 各列の内容を確認して、必要に応じて列幅を調整
    for (let col = 1; col <= numCols; col++) {
      const headerValue = headers[col - 1];
      const headerWidth = headerValue ? headerValue.toString().length * 1.2 : 10;
      
      // データ行の最大幅を確認
      let maxDataWidth = 0;
      if (numRows > 0) {
        for (let row = 2; row <= numRows + 1; row++) {
          const cellValue = sheet.getRange(row, col).getValue();
          if (cellValue) {
            const cellWidth = cellValue.toString().length * 1.1;
            if (cellWidth > maxDataWidth) {
              maxDataWidth = cellWidth;
            }
          }
        }
      }
      
      // ヘッダーとデータの大きい方に余裕を持たせて設定
      const finalWidth = Math.max(headerWidth, maxDataWidth, 10) + 2;
      sheet.setColumnWidth(col, Math.min(finalWidth, 300)); // 最大300ピクセル
    }
  } catch (e) {
    Logger.log(`> (div) ${fileName}: 列幅の自動調整に失敗しました。 ${e.message}`);
  }

  return { url: newSs.getUrl(), rowCount: numRows };
}

/**
 * (div) メイン処理2: URL一覧シートを読み込み、各シートの回答をマージします。
 */
function div_mergeResponseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName(DIV_SUMMARY_SHEET_NAME);
  if (!summarySheet) {
    SpreadsheetApp.getUi().alert(`エラー (div): 「${DIV_SUMMARY_SHEET_NAME}」が見つかりません。`);
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
      Logger.log(`(div) スキップ: ${personName} (無効なURL)`);
      continue;
    }

    try {
      const targetSs = SpreadsheetApp.openByUrl(url);
      const targetSheet = targetSs.getSheets()[0];
      const values = targetSheet.getDataRange().getValues();

      if (!headersSet) {
        headers = values.shift();
        headers.unshift('確認先');
        headers.push('参照元URL');
        mergedData.push(headers);
        headersSet = true;
      } else {
        values.shift();
      }

      values.forEach(dataRow => {
        dataRow.unshift(personName);
        dataRow.push(url);
        mergedData.push(dataRow);
      });
      Logger.log(`(div) マージ完了: ${personName}`);

    } catch (e) {
      Logger.log(`(div) エラー: ${personName} のシート読み込み失敗。 (URL: ${url}) ${e.message}`);
      if (headersSet) {
        const errorRow = new Array(headers.length).fill('');
        errorRow[0] = personName;
        errorRow[1] = `シートの読み込みに失敗しました: ${e.message}`;
        errorRow[errorRow.length - 1] = url; // 最後の列にURLを設定
        mergedData.push(errorRow);
      }
    }
  }

  if (mergedData.length <= 1) { // 1 = ヘッダーのみ
    Logger.log('(div) マージするデータがありませんでした。');
    SpreadsheetApp.getUi().alert('(div) マージするデータがありませんでした。');
    return;
  }

  const outputSheetName = 'divマージ済み回答';
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
  SpreadsheetApp.flush();
  Logger.log('(div) マージ処理が完了しました。');
  SpreadsheetApp.getUi().alert('(div) 回答シートマージ処理が完了しました。');
}