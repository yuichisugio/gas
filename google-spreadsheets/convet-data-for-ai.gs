//AI用のソース用にslide出力
function exportSpreadsheetToPDF() {
	const spreadsheetId = '1Aqk5VdPpxxx';
	const folderId = '1sxxxxxxxxxxxxxxxx';
	const sheetName = 'AI用マニュアル更新';

	const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf&portrait=false&exportFormat=pdf`;

	const token = ScriptApp.getOAuthToken();
	const headers = {
		Authorization: 'Bearer ' + token
	};

	const response = UrlFetchApp.fetch(url, { headers: headers });
	const blob = response.getBlob();

	// 生成日時を取得し、ファイル名を作成
	const now = new Date();
	const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH:mm');
	const fileName = formattedDate + 'マニュアル.PDF';
	blob.setName(fileName);

	// PDFを指定フォルダに保存
	const folder = DriveApp.getFolderById(folderId);
	const file = folder.createFile(blob);

	// ダウンロードリンクを取得
	const downloadUrl = file.getUrl();

	// スプレッドシートに書き込み
	const ss = SpreadsheetApp.openById(spreadsheetId);
	const sheet = ss.getSheetByName(sheetName);
	sheet.getRange('A2').setValue(downloadUrl);
	sheet.getRange('B2').setValue(formattedDate);

	Logger.log('PDF URL: ' + downloadUrl);
	Logger.log('生成日時: ' + now);
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

function convertSpreadsheetToDoc() {
	try {
		// スプレッドシートIDを指定
		var spreadsheetId = 'xxxxxxxxx';
		var sheetName = 'xxxxxxxxx'; // シート名を指定
		Logger.log('Spreadsheet ID: ' + spreadsheetId);

		// スプレッドシートを取得
		var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
		Logger.log('Spreadsheet opened successfully');

		var sheet = spreadsheet.getSheetByName(sheetName);
		Logger.log('Sheet opened successfully: ' + sheetName);

		// データを取得
		var data = sheet.getDataRange().getValues();
		Logger.log('Data retrieved successfully');

		// 新しいGoogleドキュメントを作成
		var doc = DocumentApp.create('Converted Spreadsheet to Doc');
		Logger.log('Document created successfully');

		var body = doc.getBody();

		// データをドキュメントに書き込む
		data.forEach(function (row) {
			var rowText = row.join('\t'); // タブ区切りで行を結合
			body.appendParagraph(rowText);
		});
		Logger.log('Data written to document successfully');

		// ドキュメントのURLを取得してログに出力
		var docUrl = doc.getUrl();
		Logger.log('Document created: ' + docUrl);
	} catch (e) {
		Logger.log('Error: ' + e.toString());
	}
}
