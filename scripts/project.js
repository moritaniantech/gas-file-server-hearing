/**
 * @fileoverview
 * (ファイル2/2) MIXI (新)project-2 シート処理用スクリプト
 * * このファイルには「project用」の処理のみが含まれます。
 * * 共通メニュー(onOpen)は Code_div.gs に記載されています。
 */

// ===================================================
// project フォルダ用 設定
// ===================================================

/** (project) 元データが記載されているシート名 */
const PROJECT_SOURCE_SHEET_NAME = '（新）project-2';

/** (project) 作成したシートのURL一覧を記載するシート名 */
const PROJECT_SUMMARY_SHEET_NAME = 'project回答シートURL一覧';

/** (project) 確認先不明の場合のファイル名 */
const PROJECT_UNKNOWN_SHEET_NAME = '確認先不明';

/** (project)【重要】回答用スプレッドシートを保存するGoogle DriveフォルダのID */
const PROJECT_DESTINATION_FOLDER_ID = '1OjgFtdJYA3kyZm95ogtiiYAkU7zLL8wM'; // ← ★★★ 設定してください ★★★

// (project) (新)project-2シートの列インデックス（0ベース）
const PROJECT_COL_A_FOLDER_NAME = 0;      // A列: projectフォルダ名
const PROJECT_COL_B_STATUS = 1;          // B列: ステータス
const PROJECT_COL_C_FOLDER_COUNT = 2;     // C列: フォルダ数
const PROJECT_COL_D_FILE_COUNT = 3;       // D列: ファイル数
const PROJECT_COL_E_DATA_SIZE = 4;        // E列: データ容量/GB
const PROJECT_COL_F_LAST_UPDATED = 5;     // F列: 最終更新日
const PROJECT_COL_G_MIGRATION_DEST = 6;   // G列: 移行先
const PROJECT_COL_H_MIGRATION_METHOD = 7; // H列: 移行方法
const PROJECT_COL_I_MANAGER = 8;          // I列: 管理者

// ===================================================
// project フォルダ用 スクリプト
// ===================================================

/**
 * (project) メイン処理1: (新)project-2シートを読み込み、確認先ごとにスプレッドシートを分割作成します。
 */
function project_createResponseSheets() {
  // ユーザー認証チェック
  const allowedUser = 'ryosuke.morita.ts@mixi.co.jp';
  const currentUser = Session.getActiveUser().getEmail();
  if (currentUser !== allowedUser) {
    SpreadsheetApp.getUi().alert('エラー (project)', `このスクリプトは ${allowedUser} のみが実行できます。\n現在のユーザー: ${currentUser}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

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

  // 「（新）project-2」シートは4行目が列名、5行目からレコードが開始
  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();
  
  // 4行目を列名として取得
  const headers = sourceSheet.getRange(4, 1, 1, lastCol).getValues()[0];
  
  // 5行目以降をデータとして取得
  let data = [];
  if (lastRow >= 5) {
    data = sourceSheet.getRange(5, 1, lastRow - 4, lastCol).getValues();
  }

  Logger.log(`(project) 元データ ${data.length}行の処理を開始します。`);
  
  // デバッグ: 最初の3行のデータを確認
  if (data.length > 0) {
    Logger.log(`(project) デバッグ: 最初の行のデータ`);
    Logger.log(`  A列(フォルダ名): "${data[0][PROJECT_COL_A_FOLDER_NAME]}" (型: ${typeof data[0][PROJECT_COL_A_FOLDER_NAME]})`);
    Logger.log(`  B列(ステータス): "${data[0][PROJECT_COL_B_STATUS]}" (型: ${typeof data[0][PROJECT_COL_B_STATUS]})`);
    Logger.log(`  I列(管理者): "${data[0][PROJECT_COL_I_MANAGER]}" (型: ${typeof data[0][PROJECT_COL_I_MANAGER]})`);
  }

  // --- (project) データをA列（projectフォルダ名）でグループ化 ---
  const dataByFolder = {};
  data.forEach((row, rowIndex) => {
    // 空行をスキップ
    if (row.join('').trim().length === 0) return;
    
    const folderName = row[PROJECT_COL_A_FOLDER_NAME];
    if (!folderName || folderName.toString().trim().length === 0) return;

    const folderNameStr = folderName.toString().trim();
    if (!dataByFolder[folderNameStr]) {
      dataByFolder[folderNameStr] = [];
    }
    dataByFolder[folderNameStr].push(row);
  });

  Logger.log(`(project) フォルダ数: ${Object.keys(dataByFolder).length}件`);

  // --- (project) 各フォルダごとに確認先を決定 ---
  // 要件: 
  // 1. G列（移行先）が「AWS S3」の場合は、確認先を「確認不要」
  // 2. I列の管理者に「×」が入力されている場合は、確認先を「確認先不明」
  // 3. I列の管理者に「×」以外が入力されている場合:
  //    - A列に同一の名称が複数行に記載されている場合、B列「ステータス」が「在職」と記入されている行から1件確認先を絞り出す
  //    - B列がすべて「在職」以外の場合、確認先不明として分類
  const folderConfirmationPerson = {};
  
  for (const folderName in dataByFolder) {
    const rows = dataByFolder[folderName];
    
    // G列（移行先）が「AWS S3」かどうかを確認（最初の行でチェック）
    const firstRow = rows[0];
    const migrationDest = firstRow[PROJECT_COL_G_MIGRATION_DEST];
    if (migrationDest && migrationDest.toString().trim() === 'AWS S3') {
      folderConfirmationPerson[folderName] = '確認不要';
      Logger.log(`(project) フォルダ「${folderName}」: G列が「AWS S3」のため、確認不要`);
      continue;
    }
    
    // B列（ステータス）が「在職」の行を探す（型変換とtrimを考慮）
    const activeRows = rows.filter(row => {
      const status = row[PROJECT_COL_B_STATUS];
      if (!status) return false;
      const statusStr = status.toString().trim();
      return statusStr === '在職';
    });
    
    Logger.log(`(project) フォルダ「${folderName}」: 全行数=${rows.length}, 在職行数=${activeRows.length}`);
    
    let confirmationPerson = PROJECT_UNKNOWN_SHEET_NAME;
    
    if (activeRows.length > 0) {
      // 「在職」の行がある場合、その中からI列（管理者）の値を確認先として選ぶ
      // I列が「×」の場合は確認先不明
      // I列が「×」以外の場合は、その値を確認先とする（最初の「在職」行のI列の値を確認先とする）
      for (let i = 0; i < activeRows.length; i++) {
        const manager = activeRows[i][PROJECT_COL_I_MANAGER];
        if (manager) {
          const managerStr = manager.toString().trim();
          Logger.log(`(project) フォルダ「${folderName}」: 在職行[${i}]のI列値="${managerStr}"`);
          if (managerStr !== '×' && managerStr.length > 0) {
            confirmationPerson = managerStr;
            Logger.log(`(project) フォルダ「${folderName}」: 確認先を「${confirmationPerson}」に決定`);
            break;
          }
        }
      }
      
      if (confirmationPerson === PROJECT_UNKNOWN_SHEET_NAME) {
        Logger.log(`(project) フォルダ「${folderName}」: 在職行のI列がすべて「×」または空のため、確認先不明`);
      }
    } else {
      // B列がすべて「在職」以外の場合、確認先不明
      Logger.log(`(project) フォルダ「${folderName}」: 在職行がないため、確認先不明`);
    }
    
    folderConfirmationPerson[folderName] = confirmationPerson;
  }

  // --- (project) 確認先ごとにデータを振り分け、他フォルダ管理者を計算 ---
  const dataByPerson = {};
  
  for (const folderName in dataByFolder) {
    const rows = dataByFolder[folderName];
    const confirmationPerson = folderConfirmationPerson[folderName];
    
    // 他フォルダ管理者の計算：
    // 「（新）project−2」のA列に同一のフォルダ名が記述されており、かつB列が「在職」のケースがあれば、
    // I列の値をすべて「,」で繋げて転記する（確認先として抽出された方の氏名は除外）
    const activeRows = rows.filter(row => {
      const status = row[PROJECT_COL_B_STATUS];
      if (!status) return false;
      const statusStr = status.toString().trim();
      return statusStr === '在職';
    });
    const otherManagers = [];
    activeRows.forEach(row => {
      const manager = row[PROJECT_COL_I_MANAGER];
      if (manager) {
        const managerStr = manager.toString().trim();
        if (managerStr !== '×' && managerStr.length > 0) {
          // 確認先として抽出された方の氏名は除外
          if (managerStr !== confirmationPerson && otherManagers.indexOf(managerStr) === -1) {
            otherManagers.push(managerStr);
          }
        }
      }
    });
    const otherManagersStr = otherManagers.join(',');

    // 各フォルダの1件のみを確認先ごとに振り分け（同一フォルダ名の重複を防ぐ）
    // 在職の行を優先し、なければ最初の行を使用
    // （activeRowsは上で既に計算済みなので再利用）
    
    // 在職の行があれば最初の1件、なければ最初の行を使用
    const selectedRow = activeRows.length > 0 ? activeRows[0] : rows[0];
    
    // 回答シートの列構成: A, C, D, E, F列 + 他フォルダ管理者 + ユーザー入力列
    const newRow = [
      selectedRow[PROJECT_COL_A_FOLDER_NAME],      // 0: projectフォルダ（A列）
      selectedRow[PROJECT_COL_C_FOLDER_COUNT],     // 1: フォルダ数（C列）
      selectedRow[PROJECT_COL_D_FILE_COUNT],       // 2: ファイル数（D列）
      selectedRow[PROJECT_COL_E_DATA_SIZE],        // 3: データ容量/GB（E列）
      selectedRow[PROJECT_COL_F_LAST_UPDATED],     // 4: 最終更新日（F列）
      otherManagersStr,                             // 5: 他フォルダ管理者
      '',                                           // 6: 回答者メールアドレス（ユーザー記入）
      selectedRow[PROJECT_COL_G_MIGRATION_DEST] || '',    // 7: 移行先（G列）
      selectedRow[PROJECT_COL_H_MIGRATION_METHOD] || '',  // 8: 移行方法（H列）
      '',                                           // 9: 共有ドライブ名（ユーザー記入）
      false,                                        // 10: 個人情報有無（チェックボックス）
      false,                                        // 11: 自動化有無（チェックボックス）
      ''                                            // 12: その他（ユーザー記入）
    ];

    if (!dataByPerson[confirmationPerson]) {
      dataByPerson[confirmationPerson] = [];
    }
    dataByPerson[confirmationPerson].push(newRow);
  }
  
  Logger.log(`(project) データ振り分け完了。確認先: ${Object.keys(dataByPerson).length}件`);
  for (const person in dataByPerson) {
    Logger.log(`(project) 確認先「${person}」: ${dataByPerson[person].length}件`);
  }

  // --- (project) 回答用シートのヘッダー定義 ---
  const outputHeaders = [
    'projectフォルダ',           // 0
    'フォルダ数',                // 1
    'ファイル数',                // 2
    'データ容量 / GB',           // 3
    '最終更新日',                // 4
    '他フォルダ管理者',          // 5
    '回答者メールアドレス',       // 6
    '移行先',                    // 7
    '移行方法',                  // 8
    '共有ドライブ名',            // 9
    '個人情報有無',              // 10
    '自動化有無',                // 11
    'その他'                     // 12
  ];
  
  // ユーザーが回答を入力する列のタイトル（色を変えるため）
  const userInputHeaderIndices = [6, 7, 8, 9, 10, 11, 12]; // 回答者メールアドレス、移行先、移行方法、共有ドライブ名、個人情報有無、自動化有無、その他

  // --- (project) データバリデーションルール ---
  const rules = {
    migrationRule: SpreadsheetApp.newDataValidation().requireValueInList(['Googleドライブ', 'Googleドライブ_別テナント', 'AWS S3', '不要'], true).build(),
    methodRule: SpreadsheetApp.newDataValidation().requireValueInList(['自対応', 'hatakan依頼', '不要'], true).build()
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
    const fileName = `(project) ${person}`;
    const rows = dataByPerson[person];
    Logger.log(`(project) シート作成開始: ${fileName} (${rows.length}件)`);
    try {
      const result = project_createAndFormatSheet(fileName, outputHeaders, rows, destinationFolder, userInputHeaderIndices, rules);
      summaryData.push([person, result.url, result.rowCount]);
      Logger.log(`(project) 作成完了: ${person} (URL: ${result.url})`);
    } catch (e) {
      Logger.log(`(project) エラー: ${person} のシート作成に失敗しました。 ${e.message}`);
      summaryData.push([person, `作成失敗: ${e.message}`, 0]);
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

  // 不要な行と列を削除
  if (sheet.getMaxRows() > numRows + 1) {
    sheet.deleteRows(numRows + 2, sheet.getMaxRows() - (numRows + 1));
  }
  if (sheet.getMaxColumns() > numCols) {
    sheet.deleteColumns(numCols + 1, sheet.getMaxColumns() - numCols);
  }

  // ヘッダー行の設定
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setValues([headers]);
  headerRange.setBackground('#d9d9d9').setFontWeight('bold').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // ユーザーが回答を入力する列のタイトルを他のタイトルと色を変える
  highlightIndices.forEach(colIndex => {
    sheet.getRange(1, colIndex + 1).setBackground('#c9daf8');
  });

  if (numRows > 0) {
    Logger.log(`> (project) ${fileName}: ${numRows}件のデータを書き込みます。`);
    try {
      const dataRange = sheet.getRange(2, 1, numRows, numCols);
      dataRange.setValues(dataRows);
      dataRange.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

      // データバリデーション設定
      // 移行先（8列目 = インデックス7）
      sheet.getRange(2, 8, numRows, 1).setDataValidation(rules.migrationRule);
      // 移行方法（9列目 = インデックス8）
      sheet.getRange(2, 9, numRows, 1).setDataValidation(rules.methodRule);
      
      // チェックボックス設定
      // 個人情報有無（11列目 = インデックス10）
      sheet.getRange(2, 11, numRows, 1).insertCheckboxes();
      // 自動化有無（12列目 = インデックス11）
      sheet.getRange(2, 12, numRows, 1).insertCheckboxes();

    } catch (e) {
      Logger.log(`> (project) ${fileName}: データ書き込みまたはフォーマット中にエラー: ${e.message}`);
      throw e;
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
    if (!url || !url.toString().startsWith('http')) {
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
        headers.push('参照元URL');
        mergedData.push(headers);
        headersSet = true;
      } else {
        values.shift(); // ヘッダー行をスキップ
      }

      values.forEach(dataRow => {
        dataRow.unshift(personName);
        dataRow.push(url);
        mergedData.push(dataRow);
      });
      Logger.log(`(project) マージ完了: ${personName}`);

    } catch (e) {
      Logger.log(`(project) エラー: ${personName} のシート読み込み失敗。 (URL: ${url}) ${e.message}`);
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
  SpreadsheetApp.flush();
  Logger.log('(project) マージ処理が完了しました。');
  SpreadsheetApp.getUi().alert('(project) 回答シートマージ処理が完了しました。');
}
