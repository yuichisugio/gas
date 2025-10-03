function tiktokAPI() {
	const sheet = SpreadsheetApp.getActiveSheet();
	const url = sheet.getRange('H6').getValue();
	let response = UrlFetchApp.fetch('https://www.tiktok.com/oembed?url=' + url);
	let json = JSON.parse(response.getContentText());

	let title = json['title'];
	let account = json['author_name'];
	let thumnail = json['thumbnail_url'];

	sheet.insertImage(thumnail, 2, 13).setWidth(300).setHeight(500);
	sheet.getRange('H4').setValue(account);
	sheet.getRange('G13').setValue(title);
}
