/**
 * @param {null}  - なし
 * @return {null} - なし。「ルールの更新内容」「フェーズ別注意事項」で更新があった場合に、更新日を入れる関数
 */
function insertLastModified() {

  // 各シートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  //getActiveCell()を使う際のポイントは、シートの取得の際もgetActive〇〇()を使用する必要がある
  const active_sheet = ss.getActiveSheet();
  const functionListSheet = ss.getSheetByName("ルールの更新内容");
  const phaseSheet = ss.getSheetByName("フェーズ別注意事項");

  // 編集されたシートが、「ルールの更新内容」または、「フェーズ別注意事項」の場合。
  if (active_sheet.getName() === functionListSheet.getName() || active_sheet.getName() === phaseSheet.getName()) {

    // ヘッダーに「更新日」が記載されているインデックス番号を探す
    let column_modified = findColumn(active_sheet, '更新日');

    // 更新したセルを取得。getActiveCell()を使う際のポイントは、シートの取得の際もgetActive〇〇()を使用する必要がある
    let active_cell = active_sheet.getActiveCell();

    // 更新した行を取得
    let active_row = active_cell.getRow();

    // headerの「更新日」の文言が修正されても、それが内容の更新だと認識されないようにしている
    if (active_row == 1) {
      return;
    }

    // 削除時は動作しないように。
    if (active_cell.getValue() != "") {

      // 更新時刻を記入
      active_sheet.getRange(active_row, column_modified + 1).setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'));
    }
  }
}


// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー


/**
 * 与えられた引数のシートの、キーワードが入っている列を探す。
 * @param {string} sheet - 列番号を取得したい項目のあるシート
 * @param {string} keyword - 取得したい項目名
 * @return {double?} target_column - keywordの列のインデックス
 */
function findColumn(sheet, keyword) {

  // 最終列を取得
  let last_column = sheet.getLastColumn();

  // ヘッダー行をすべて（最終列まで）取得
  let header = sheet.getRange(1, 1, 1, last_column).getValues()[0];

  // 特定の文字列のある列がどこかを検索して、列のインデックスを返す。シートの列番号として処理したい場合は +1 することに注意。
  let target_column = header.indexOf(keyword);

  // 「keyword」のインデックス番号を返す
  return target_column;
}
