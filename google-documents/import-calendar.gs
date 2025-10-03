/**
 * 日報などで、ドキュメントにカレンダーの予定を入れる
 */

function onOpen() {
	// メニューにボタンを追加
	const menu = DocumentApp.getUi();
	menu
		.createMenu('テンプレ追加')
		.addItem('ver1', 'insert_template_ver1')
		.addItem('ver2', 'insert_template_ver2')
		.addItem('ver3', 'insert_template_ver3')
		.addToUi();
}

// 初めのバージョン！
function insert_template_ver1() {
	// 文章を挿入する
	const body = DocumentApp.getActiveDocument().getBody();

	body.insertParagraph(0, '\n\n\n\n\n\n\n').setLineSpacing(1.5);
	body.insertParagraph(0, '「感想」').setLineSpacing(1.5).editAsText().setFontSize(12).setBold(true);
	body.insertListItem(0, '日報の作成').setGlyphType(DocumentApp.GlyphType.BULLET).setLineSpacing(1.5);
	for (let i = 0; i < 6; i++) {
		body.insertListItem(0, '').setGlyphType(DocumentApp.GlyphType.BULLET).setLineSpacing(1.5);
	}
	body.insertListItem(0, '会議').setGlyphType(DocumentApp.GlyphType.BULLET).setLineSpacing(1.5);
	body.insertListItem(0, 'Googleカレンダーに予定を反映').setGlyphType(DocumentApp.GlyphType.BULLET).setLineSpacing(1.5);
	body.insertListItem(0, 'Slack確認&返信').setGlyphType(DocumentApp.GlyphType.BULLET).setLineSpacing(1.5);
	body.insertListItem(0, 'freee勤怠管理で出勤退勤').setGlyphType(DocumentApp.GlyphType.BULLET).setLineSpacing(1.5);
	body.insertParagraph(0, '「今日の行動」').setLineSpacing(1.5).editAsText().setFontSize(12).setBold(true);
	body.insertParagraph(0, getYYYYMMDDW()).setLineSpacing(1.5).editAsText().setFontSize(16).setBold(true);
}

// 曜日によって、入れる内容を切り替えられるバージョン！
function insert_template_ver2() {
	// 文章を挿入する
	const body = DocumentApp.getActiveDocument().getBody();
	const today = new Date();
	const weekdaylist = ['日', '月', '火', '水', '木', '金', '土'];
	var w = weekdaylist[today.getDay()];
	var yyyyMMdd = Utilities.formatDate(today, 'JST', 'yyyy年MM月dd日') + '(' + w + ')';
	let task = [
		'日報の作成',
		'会議',
		'',
		'Slack確認&返信',
		'タスク管理&カレンダーに予定を反映',
		'freee勤怠管理で出勤退勤'
	];

	body.insertParagraph(0, '\n\n\n\n\n').setLineSpacing(1.5);
	body.insertParagraph(0, '「感想」').setLineSpacing(1.5).editAsText().setFontSize(12).setBold(true);

	for (let i = 0; i < task.length; i++) {
		if (task[i] == '') {
			for (let s = 0; s < 6; s++) {
				body.insertListItem(0, '').setGlyphType(DocumentApp.GlyphType.BULLET).setLineSpacing(1.5);
			}
			if (w == '月') {
				body.insertListItem(0, '定例').setGlyphType(DocumentApp.GlyphType.BULLET).setLineSpacing(1.5);
				body.insertListItem(0, '定例-1').setGlyphType(DocumentApp.GlyphType.BULLET).setLineSpacing(1.5);
			} else if (w == '火') {
				body.insertListItem(0, '定例-2').setGlyphType(DocumentApp.GlyphType.BULLET).setLineSpacing(1.5);
			} else if (w == '木') {
				body.insertListItem(0, '定例-3').setGlyphType(DocumentApp.GlyphType.BULLET).setLineSpacing(1.5);
			}
		} else {
			body.insertListItem(0, task[i]).setGlyphType(DocumentApp.GlyphType.BULLET).setLineSpacing(1.5);
		}
	}

	body.insertParagraph(0, '「今日の行動」').setLineSpacing(1.5).editAsText().setFontSize(12).setBold(true);
	body.insertParagraph(0, yyyyMMdd).setLineSpacing(1.5).editAsText().setFontSize(16).setBold(true);
}

// Googleカレンダーの予定も入れてくれるバージョン！
function insert_template_ver3() {
	// 今日の日付を取得
	const today = new Date();

	// 週のリストを作成
	const weekday_list = ['日', '月', '火', '水', '木', '金', '土'];

	// 今日の曜日を取得。getDay()はインデックス番号を返す
	var weekday = weekday_list[today.getDay()];

	// 日付の出力を整える
	var yyyyMMdd = Utilities.formatDate(today, 'JST', 'yyyy年MM月dd日') + '(' + weekday + ')';

	// 出したいタスクリストを用意
	let task_list = get_today_calender_task_title();

	// 前日の記載に改行を持たせる
	insert_text('para', '\n\n\n\n');

	// 感想を入れる
	insert_text('para', '感想');

	// カレンダーの予定を入れる
	for (let i = 0; i < task_list.length; i++) {
		insert_text('list', task_list[i]);
	}

	// 文言を入れる
	insert_text('para', '今日のタスク');

	// 今日の日付を入れる
	insert_text('para', yyyyMMdd);
}

// ドキュメントに指定のテキストを入れている
function insert_text(type, title) {
	// ドキュメントを取得
	const body = DocumentApp.getActiveDocument().getBody();

	// 日付の正規表現
	const day_regex = /^[0-2]{2}[0-9]{2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日\([日月火水木金土]\)/gi;

	// type別に、入れる形式を変えている
	if (type == 'para') {
		// 感想やタスクの場合は大文字
		if (title == '感想' || title == '今日のタスク') {
			body.insertParagraph(0, title).setLineSpacing(1.5).editAsText().setFontSize(12).setBold(true);

			// 日付の場合を正規表現でとる
		} else if (day_regex.test(title)) {
			body.insertParagraph(0, title).setLineSpacing(1.5).editAsText().setFontSize(16).setBold(true);

			// それ以外の内容
		} else {
			body.insertParagraph(0, title).setLineSpacing(1.5);
		}

		// リストに入れるタスク。カレンダーの内容が入る
	} else if (type == 'list') {
		body.insertListItem(0, title).setGlyphType(DocumentApp.GlyphType.BULLET).setLineSpacing(1.5);
	}
}

// 今日のカレンダーの予定のタイトル一覧を配列で返す。
function get_today_calender_task_title() {
	// カレンダーを取得
	const calender = CalendarApp.getDefaultCalendar();

	// 今日の予定を取得
	const events = calender.getEventsForDay(new Date());

	// 今日の予定のタイトルが入る配列
	let task_array = [];

	// 今日の予定から配列を抜き出す処理
	for (let i = 0; i < events.length; i++) {
		// 終日イベント（自宅勤務）などのタイトルはスキップ
		if (events[i].isAllDayEvent()) {
			continue;
		}
		// 昼休憩は、日報に不要なのでスキップ
		if (events[i].getTitle() == '昼休憩') {
			continue;
		}
		// 上記以外を入れる
		task_array.push(events[i].getTitle());
	}

	// 配列の順番を逆転させて、最後の予定からドキュメントに入れられるようにしている。
	task_array.reverse();

	// 配列を呼び出し元に返している
	return task_array;
}
