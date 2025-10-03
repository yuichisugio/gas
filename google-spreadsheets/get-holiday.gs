/*
目的：
この先5年分くらいの祝日の一覧を作成して、それを↓シートの祝日リストなどに貼り付けておきたい
さらに、営業日計算シートにも

要件：
・祭日を抜いたバージョンが欲しい
・リマインド、スケジュール自動作成シートのGASの祝日リスト、営業日計算シートの祝日タブ、の3つに入れる
・「yyyy-MM-dd」の形式

祝日のみ取得できるカレンダーID
ja.japanese.official#holiday@group.v.calendar.google.com
*/

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

/*
営業日計算シートの祝日タブ用の関数

↓みたいな結果が出る
2022/01/01	元日
2022/01/10	成人の日
2022/02/11	建国記念の日
2022/02/23	天皇誕生日
*/
function setHolidayAndNameList() {
	// シートを取得
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('作業用');

	// 前回の記載内容を削除
	sheet.clearContents();

	// 祭日を含まない祝日を取得するカレンダーIDを指定
	const holidayCalendar = CalendarApp.getCalendarById('ja.japanese.official#holiday@group.v.calendar.google.com');

	// カレンダーの開始日を指定
	const startDate = new Date('2024-01-01');

	// カレンダーの終了日を指定。↑のカレンダーIDは6年しか取得できないらしいので、24年からだと30年まで取れない（29年まで）
	const endDate = new Date('2030-12-31');

	// 開始と終了を指定して、祝日を取得
	const holidayEvents = holidayCalendar.getEvents(startDate, endDate);

	// 重複を削除し、日付と祝日名を取得する
	let holidayList = holidayEvents.map((event) => {
		// 日付
		let date = Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
		// 祝日の名前
		let title = event.getTitle();
		// 日付と名前を配列として返す。date と title をオブジェクト { date, title } で返すように変更しました。このオブジェクトを使って日付をキーとして重複を判定します。
		// オブジェクトにした方が、配列でインデックス番号で指定するより分かりやすい
		return { date, title };
	});

	// 日付ベースで重複を削除する。
	// reduce() 関数を使い、holidayList の中から日付が同じものを1つにまとめる処理を追加しました。
	// acc は蓄積用の配列です。find() で既に同じ日付が存在するかを確認し、存在しない場合だけ acc に追加します。これにより、同じ日付の祝日は1つにまとめられます。
	// item変数は、acc配列の中の要素。それがreduceの今回の要素と同じか見ている
	let uniqueHolidays = holidayList.reduce((acc, current) => {
		const x = acc.find((item) => item.date === current.date);
		// 違う場合は配列に追加して、同じ場合はこのターンのreduceをスキップしている
		if (!x) {
			acc.push(current);
		}
		return acc;
	}, []);

	// 2次元配列に変換してシートに出力する。重複を削除した祝日リストを2次元配列に変換し、シートに出力できる形式にします。
	let outputData = uniqueHolidays.map((holiday) => [holiday.date, holiday.title]);

	// シートに出力する
	sheet.getRange(2, 1, outputData.length, 2).setValues(outputData);
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

/*
OPEリマインド、スケジュール自動作成シートのGASの祝日リスト用の祝日リスト作成関数

↓みたいな結果が出る
"2024-01-01","2024-01-02","2024-01-03","2024-01-08","2024-02-11","2024-02-12","2024-02-23","2024-03-20",
"2024-04-29","2024-05-03","2024-05-04","2024-05-05","2024-05-06","2024-07-15","2024-08-11","2024-08-12",
"2024-09-16","2024-09-22","2024-09-23","2024-10-14","2024-11-03","2024-11-04","2024-11-23","2024-12-31",
"2025-01-01","2025-01-02","2025-01-03","2025-01-13","2025-02-11","2025-02-23","2025-02-24","2025-03-20"
*/
function setHolidayList() {
	// シートを取得
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('作業用');

	// 前回の記載内容を削除
	sheet.clearContents();

	// 祭日を含まない祝日を取得するカレンダーIDを指定
	const holidayCalendar = CalendarApp.getCalendarById('ja.japanese.official#holiday@group.v.calendar.google.com');

	// カレンダーの開始日を指定
	const startDate = new Date('2024-01-01');

	// カレンダーの終了日を指定。↑のカレンダーIDは6年しか取得できないらしいので、24年からだと30年まで取れない（29年まで）
	const endDate = new Date('2030-12-31');

	// 開始と終了を指定して、祝日を取得
	const holidayEvents = holidayCalendar.getEvents(startDate, endDate);

	// map関数で、配列を一つづく取り出して、形式を整えて、結果を再度配列に入れていく
	let holidayList = holidayEvents.map((event) =>
		Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
	);

	// 元旦など祝日が重複している日の重複を削除する。Setオブジェクトは重複する値を自動的に排除する特性を持っているので、holidayListをSetに変換し、再び配列に戻すことで重複を取り除きます。
	holidayList = [...new Set(holidayList)]; // Setに変換して重複を削除Ï

	// 8個ごとに改行する
	let output = '';
	for (let i = 0; i < holidayList.length; i++) {
		// 日付の初めと最後に"を入れて、最後にカンマを追加する
		output += '"' + holidayList[i] + '",';
		if ((i + 1) % 8 == 0) {
			output += '\n';
		}
	}

	// 最後のカンマを削除して出力
	output = output.trim().replace(/,$/, '');

	// 1つのセルに日付リストを出力
	sheet.getRange(2, 1).setValue(output);
}
