// 使用するシートを取得
const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
const makeFolderSheet = spreadSheet.getSheetByName('フォルダ作成');
const campaignSheet = spreadSheet.getSheetByName('キャンペーン一覧');
const scheduleParseSheet = spreadSheet.getSheetByName('スプシのスケ→teamXxxx一覧');
const slackSheet = spreadSheet.getSheetByName('チャンネル設定');

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// クライアントフォルダのIDと変数を作成
const clientFolderId = '1jo-xxxx-xxxxxxxx';
const clientFolder = DriveApp.getFolderById(clientFolderId);

// ブランド分析のテンプレシートのIDと変数を作成
const brandAnalysis_templateSheetId = '1Bxxxxxxxxxxxxxxxxxxxxxxxxxx';
const brandAnalysis_templateSheet = DriveApp.getFileById(brandAnalysis_templateSheetId);

// ブランド分析を格納するフォルダのIDと変数を作成
let brandAnalysis_FolderId = '19aIxxxxxxxxxxxxxxxxxxxxxxxxx';
let brandAnalysis_Folder = DriveApp.getFolderById(brandAnalysis_FolderId);

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// slack_bot(teamXxxx_bot)のtoken。コードのコピペで流出しないようにスクリプトプロパティに入れているので、そこから取り出す
const token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// TeamXxxxFlagRangeクラス
class TeamXxxxFlagRange {
	constructor(sheet, headers, row, columnName) {
		this.flag = 0;
		this.columnNumber = headers[columnName];
		if (!this.columnNumber) {
			throw new Error(`ヘッダー名 "${columnName}" が見つかりません。`);
		}
		this.range = sheet.getRange(row, this.columnNumber);
	}
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 引数で受け取ったテキストをslackに飛ばす関数。
function sendSlack(text, team) {
	// CPがない場合は、リンクを入れようとしたら、<|CP>のみ残ってしまうので、「CP未作成」に置き換えている
	if (text.includes('<|CP>')) {
		text = text.replace(/<\|CP>/g, 'CP未作成');
	}

	// スコープ外で宣言
	let webhookUrl = '';
	// slackのリマインドの名前
	let username = '';

	// opeチャンネル or pmチャンネルに送るか使い分けられるようにしている。
	if (team == 'ope') {
		// Slack「#ch_placeholder_ope」チャンネルのIncoming Webhook URL。↓でincoming webhooksが削除予定のため、slack appを作成して、そこから流す予定
		username = 'OPEチーム向けリマインド！';

		// Slack「#ch_placeholder_ope」チャンネルのIncoming Webhook URL。
		webhookUrl = 'https://hooks.slack.com/services/XXXXXXXXX/XXXXX/XXXXXXXX';

		// テスト用。teamXxxxテストチャンネルのIncoming Webhook URL
		// webhookUrl = "https://hooks.slack.com/services/XXXXXXXXX/XXXXX/XXXXXXXX"
	} else if (team == 'pm') {
		// Slack「#ch_placeholder_pm」チャンネルのIncoming Webhook URL。↓でincoming webhooksが削除予定のため、slack appを作成して、そこから流す予定
		username = 'PMチーム向けリマインド！';

		// Slack「#ch_placeholder_pm」チャンネルのIncoming Webhook URL。
		webhookUrl = 'https://hooks.slack.com/services/XXXXXXXXX/XXXXX/XXXXXXXX';

		// テスト用。teamXxxxテストチャンネルのIncoming Webhook URL
		// webhookUrl = "https://hooks.slack.com/services/XXXXXXXXX/XXXXX/XXXXXXXX"
	} else if (team == 'teamXxxx') {
		username = 'teamXxxx向けリマインド！';
		// teamXxxxDMのIncoming Webhook URL
		webhookUrl = 'https://hooks.slack.com/services/XXXXXXXXX/XXXXX/XXXXXXXX';
	} else if (team == 'kol') {
		// Slack「#ch_placeholder_kol」チャンネルのIncoming Webhook URL
		webhookUrl = 'https://hooks.slack.com/services/XXXXXXXXX/XXXXX/XXXXXXXX';
		username = 'KOLチーム向けリマインド！';

		// テスト用。teamXxxxテストチャンネルのIncoming Webhook URL
		// webhookUrl = "https://hooks.slack.com/services/XXXXXXXXX/XXXXX/XXXXXXXX"
	}

	// slackのアイコン
	const icon = ':pencil:';

	// Slackに送るjsonデータを作成
	const jsonData = {
		username: username,
		icon_emoji: icon,
		text: text
	};

	// jsonデータを文字列化する
	const payload = JSON.stringify(jsonData);

	// Slackに送るデータと通信方法を入れたopitionをjsonで設定
	const options = {
		method: 'post',
		contentType: 'application/json',
		payload: payload
	};

	// webhook URLにオプションの内容を飛ばす
	UrlFetchApp.fetch(webhookUrl, options);

	return 'DONE';
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 引数で、①探したいシート、②ヘッダーが何行目か指定。
// 戻り値で、①キーがヘッダーの各列名、②値がその列が何列目か、が入ったオブジェクトを返す。
function getHeaders(sheet, headersRow) {
	const headerRow = sheet.getRange(headersRow, 1, 1, sheet.getLastColumn()).getValues()[0];
	const headerMap = {};
	headerRow.forEach((header, index) => {
		headerMap[header] = index + 1;
	});
	return headerMap;
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 祝日判定関数。引数が祝日の場合trueを返す。GoogleカレンダーIDで祝日取得しても良さげ。
function isHoliday(date) {
	// 祝日リスト
	const holidays = [
		'2024-01-01',
		'2024-01-02',
		'2024-01-03',
		'2024-01-08',
		'2024-02-11',
		'2024-02-12',
		'2024-02-23',
		'2024-03-20',
		'2024-04-29',
		'2024-05-03',
		'2024-05-04',
		'2024-05-05',
		'2024-05-06',
		'2024-07-15',
		'2024-08-11',
		'2024-08-12',
		'2024-09-16',
		'2024-09-22',
		'2024-09-23',
		'2024-10-14',
		'2024-11-03',
		'2024-11-04',
		'2024-11-23',
		'2024-12-31',
		'2025-01-01',
		'2025-01-02',
		'2025-01-03',
		'2025-01-13',
		'2025-02-11',
		'2025-02-23',
		'2025-02-24',
		'2025-03-20',
		'2025-04-29',
		'2025-05-03',
		'2025-05-04',
		'2025-05-05',
		'2025-05-06',
		'2025-07-21',
		'2025-08-11',
		'2025-09-15',
		'2025-09-23',
		'2025-10-13',
		'2025-11-03',
		'2025-11-23',
		'2025-11-24',
		'2025-12-31',
		'2026-01-01',
		'2026-01-02',
		'2026-01-03',
		'2026-01-12',
		'2026-02-11',
		'2026-02-23',
		'2026-03-20',
		'2026-04-29',
		'2026-05-03',
		'2026-05-04',
		'2026-05-05',
		'2026-05-06',
		'2026-07-20',
		'2026-08-11',
		'2026-09-21',
		'2026-09-22',
		'2026-09-23',
		'2026-10-12',
		'2026-11-03',
		'2026-11-23',
		'2026-12-31',
		'2027-01-01',
		'2027-01-02',
		'2027-01-03',
		'2027-01-11',
		'2027-02-11',
		'2027-02-23',
		'2027-03-21',
		'2027-03-22',
		'2027-04-29',
		'2027-05-03',
		'2027-05-04',
		'2027-05-05',
		'2027-07-19',
		'2027-08-11',
		'2027-09-20',
		'2027-09-23',
		'2027-10-11',
		'2027-11-03',
		'2027-11-23',
		'2027-12-31',
		'2028-01-01',
		'2028-01-02',
		'2028-01-03',
		'2028-01-10',
		'2028-02-11',
		'2028-02-23',
		'2028-03-20',
		'2028-04-29',
		'2028-05-03',
		'2028-05-04',
		'2028-05-05',
		'2028-07-17',
		'2028-08-11',
		'2028-09-18',
		'2028-09-22',
		'2028-10-09',
		'2028-11-03',
		'2028-11-23',
		'2028-12-31',
		'2029-01-01',
		'2029-01-02',
		'2029-01-03',
		'2029-01-08',
		'2029-02-11',
		'2029-02-12',
		'2029-02-23',
		'2029-03-20',
		'2029-04-29',
		'2029-04-30',
		'2029-05-03',
		'2029-05-04',
		'2029-05-05',
		'2029-07-16',
		'2029-08-11',
		'2029-09-17',
		'2029-09-23',
		'2029-09-24',
		'2029-10-08',
		'2029-11-03',
		'2029-11-23',
		'2029-12-31',
		'2024-12-29',
		'2024-12-30',
		'2025-12-29',
		'2025-12-30',
		'2026-12-29',
		'2026-12-30'
	];

	// 引数に渡された日が祝日リストに含まれているならtrueを返す関数
	return holidays.includes(Utilities.formatDate(new Date(date), 'JST', 'yyyy-MM-dd'));
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 次の営業日を取得する関数。countに入れた数だけ次の営業日を返す。
function getNextBusinessDay(startDate, count) {
	// 引数の日付を複製
	let nextDate = startDate ? new Date(startDate) : new Date();

	// countの数だけ繰り返して営業日を先に進める
	for (let i = 0; i < count; i++) {
		do {
			nextDate.setDate(nextDate.getDate() + 1);
		} while (nextDate.getDay() === 0 || nextDate.getDay() === 6 || isHoliday(nextDate));
	}

	return nextDate;
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 前の営業日を取得する関数。countに入れた数だけ前の営業日を返す。
function getPreviousBusinessDay(startDate, count) {
	// 引数の日付を複製
	let previousDate = startDate ? new Date(startDate) : new Date();

	// countの数だけ繰り返して営業日を先に進める
	for (let i = 0; i < count; i++) {
		do {
			previousDate.setDate(previousDate.getDate() - 1);
		} while (previousDate.getDay() === 0 || previousDate.getDay() === 6 || isHoliday(previousDate));
	}

	return previousDate;
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// テスト
function test12() {
	// console.log(Utilities.formatDate(getPreviousBusinessDay(new Date(), 3), "JST", "yyyy/MM/dd"));
	console.log(Utilities.formatDate(getBusinessDay(new Date(2024, 8, 28)), 'JST', 'yyyy/MM/dd'));
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

const YAHOO_CLIENT_ID = 'dj00aiZpPXJHVEFMMTFwN0tqTiZzPWNvbnN1bWVyc2VjcmV0Jng9YWE-';
const YAHOO_KW_SEARCH_API_URL = 'https://jlp.yahooapis.jp/MAService/V2/parse?appid=' + YAHOO_CLIENT_ID;

//---------------------------------------------------------------------------------------------------------------------

const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
const input_sheet = spreadSheet.getSheetByName('text_input');
const output_sheet = spreadSheet.getSheetByName('text_output');
const makeFolderSheet = spreadSheet.getSheetByName('フォルダ作成');

//---------------------------------------------------------------------------------------------------------------------

// クライアントフォルダのIDと変数
const clientFolderId = '1jo-xxxx-xxxxxxxx';
const clientFolder = DriveApp.getFolderById(clientFolderId);

//---------------------------------------------------------------------------------------------------------------------

// ブランド分析シートのテンプレIDと変数
const brandAnalysis_templateSheetId = '1Bxxxxxxxxxxxxxxxxxxxxxxxxxx';
const brandAnalysis_templateSheet = DriveApp.getFileById(brandAnalysis_templateSheetId);

//---------------------------------------------------------------------------------------------------------------------

// ブランド分析フォルダIDと変数
let brandAnalysis_FolderId = '19aIxxxxxxxxxxxxxxxxxxxxxxxxx';
let brandAnalysis_Folder = DriveApp.getFolderById(brandAnalysis_FolderId);

//---------------------------------------------------------------------------------------------------------------------

// PR判定・アフィ判定用のbooleanを作成
function judge_pr_afl(type, array, sns) {
	// Instagramだけど、アフィ配列で判断させないために必要？Instagramだったらアフィ判定が必ずfalseになるためにもtypeは必要
	if (type == 'afl' && array !== null && sns == 'X(Twitter)') {
		return true;
	} else if (type == 'pr' && array !== null) {
		return true;
	} else {
		return false;
	}
}

//---------------------------------------------------------------------------------------------------------------------

// 引数の正規表現に、引数のキャプションがあるか確認
function get_regexPattern_array(regexPattern, caption) {
	return caption.match(regexPattern);
}

//---------------------------------------------------------------------------------------------------------------------

// PRとアフィの正規表現を呼び出す
function get_regexPattern(type, dialog_pr_kw) {
	let pr_regexPattern =
		dialog_pr_kw +
		'|keyword_xxxx01|keyword_xxxx02|keyword_xxxx03|keyword_xxxx04|keyword_xxxx05|keyword_xxxx06|keyword_xxxx07|keyword_xxxx08|keyword_xxxx09|keyword_xxxx10|keyword_xxxx11|keyword_xxxx12|keyword_xxxx13|keyword_xxxx14|keyword_xxxx15|keyword_xxxx16|keyword_xxxx17|keyword_xxxx18|keyword_xxxx19|keyword_xxxx20|keyword_xxxx21|keyword_xxxx22|keyword_xxxx23|keyword_xxxx24|keyword_xxxx25';

	let pr_regexPattern_reg = new RegExp(pr_regexPattern, 'ig');

	// アフィ判定KW
	let afl_regexPattern = 'affiliate_keyword_01|affiliate_keyword_02|affiliate_keyword_03|affiliate_keyword_04|affiliate_keyword_05|affiliate_keyword_06|affiliate_keyword_07|affiliate_keyword_08';

	let afl_regexPattern_reg = new RegExp(afl_regexPattern, 'ig');

	// ハッシュタグ・メンション
	// 文字列内に入れたら、\が一つ消されるので、エスケープするため\を2個入れている
	let hashtag_regexPattern = '[@#][^\\s^）@#]+';

	let hashtag_regexPattern_reg = new RegExp(hashtag_regexPattern, 'ig');

	// 全部のせ
	let all_regexPattern_reg = new RegExp(hashtag_regexPattern + '|' + pr_regexPattern + '|' + afl_regexPattern, 'ig');

	if (type == 'pr') {
		return pr_regexPattern_reg;
	} else if (type == 'afl') {
		return afl_regexPattern_reg;
	} else if (type == 'hash') {
		return hashtag_regexPattern_reg;
	} else if (type == 'all') {
		return all_regexPattern_reg;
	}
}

//---------------------------------------------------------------------------------------------------------------------

// 配列に値がないときに、その旨を記載した値を入れる関数
function isNull_and_getMessage(array) {
	let null_message = '該当する投稿なし！';

	if (array.length !== 0) {
		return;
	} else {
		array.push(null_message, -1);
	}
}

//---------------------------------------------------------------------------------------------------------------------

// 品詞テスト。単語と品詞のリストを作成して、ログに出す
function hinshi_test() {
	// キャプションの列の最終行を取得している。getlastrow()だと、投稿SNS検知の「投稿なし」に反応して無駄に取得してしまう
	let caption_lastRow = get_designation_last_row(2);

	// 2行目、１列目、データがある行まで、３列目まで取得
	let kw_captions_array = input_sheet.getRange(2, 1, caption_lastRow, 3).getValues();

	let result_array = [];

	// 1つの投稿づつYAHOO APIに投げている
	for (let j = 0; j < kw_captions_array.length; j++) {
		// キャプション
		let kw_caption = kw_captions_array[j][1];

		// 確認済みの投稿件数を出している
		console.log('確認済みの投稿件数：' + j);

		// APIを叩いている
		let yahoo_api_response_json = yahooApiRequest(kw_caption);

		let yahoo_api_response_object = JSON.parse(yahoo_api_response_json);

		let response_tokens_array = yahoo_api_response_object['result'].tokens;

		// response_tokens_arrayから、単語だけを取り出している
		for (let s = 0; s < response_tokens_array.length; s++) {
			let target_word = response_tokens_array[s][0];
			let hinshi = response_tokens_array[s][3];

			result_array.push([target_word, hinshi]);
		}
	}
	console.log(result_array);
}

//---------------------------------------------------------------------------------------------------------------------

// 形態素解析のAPIに投げている
function yahooApiRequest(queryText) {
	// リクエストヘッダ
	const headers = {
		'Content-Type': 'application/json'
	};

	// リクエストボディ
	const payload = {
		id: '1234-1',
		jsonrpc: '2.0',
		method: 'jlp.maservice.parse',
		params: {
			q: queryText
		}
	};

	// fetchのパラメータ
	const options = {
		headers: headers,
		payload: JSON.stringify(payload)
	};

	// Yahoo!形態素解析API実行
	let response_json = UrlFetchApp.fetch(YAHOO_KW_SEARCH_API_URL, options);

	// JSONをパースして、オブジェクトにしている
	let response_object = JSON.parse(response_json);

	// APIでエラーが返ってきたら、エラー文言を返す
	for (let i of Object.keys(response_object)) {
		if (i == 'error') {
			return response_object['error'].message;
		}
	}

	// tokensの中身だけを取得
	let response_tokens_array = response_object['result'].tokens;

	return response_tokens_array;
}

//---------------------------------------------------------------------------------------------------------------------

function api_test() {
	let response = {
		id: null,
		jsonrpc: '2.0',
		error: {
			code: -32700,
			message: 'Parse error'
		}
	};

	// errorメッセージがあるか確認している
	for (let i of Object.keys(response)) {
		console.log(i);
		if (i == 'error') {
			console.log(response['error'].message);
		}
	}
}

//---------------------------------------------------------------------------------------------------------------------

// 追加PR_KWを聞くダイアログを表示
function get_aditional_pr_kw() {
	let dialog_message =
		'PR投稿特有のKWがあれば入れてね！\\n「ガチモニター_商品名」「ブランド名_ad」「ReFaタイム」など\\n複数入れたい場合は、「、」で区切ってね！\\n例)レチノール、ゴリラ\\n※注意：enterを押すと文言確定ではなく送信されるので確定せずに、「、」を押す。\\n、は全角';

	let dialog_pr_kw = Browser.inputBox(dialog_message, Browser.Buttons.YES_NO);

	if (dialog_pr_kw == 'no' || dialog_pr_kw == 'cancel' || dialog_pr_kw == '') {
		dialog_pr_kw = 'keyword_xxxx';
	}

	// わかりやすいように「、」で区切るように依頼して、こちらで、「｜」に修正している
	if (dialog_message.includes('、')) {
		dialog_pr_kw = dialog_pr_kw.replace('、', '|');
	}

	return dialog_pr_kw;
}

//---------------------------------------------------------------------------------------------------------------------

// 指定列の最後の値が入っている行数を取得する関数
function get_designation_last_row(column) {
	// 一番下の行に移動してから、command+↑を押したのと同じ動きで、行を取得している
	let designation_last_row = input_sheet
		.getRange(input_sheet.getMaxRows(), column)
		.getNextDataCell(SpreadsheetApp.Direction.UP)
		.getRow();

	return designation_last_row;
}

//---------------------------------------------------------------------------------------------------------------------

function string_code_test() {
	let word = 'レチノール';
	console.log(word.charCodeAt(0));
}

//---------------------------------------------------------------------------------------------------------------------

// 絵文字除去関数
// https://highmoon-miyabi.net/blog/2022/02/21_000588.html#comments
function removeEmoji(in_value) {
	// 大体の絵文字の文字コードを入れている
	const ranges = [
		'[\ud800-\ud8ff][\ud000-\udfff]', // 基本的な絵文字除去
		'[\ud000-\udfff]{2,}', // サロゲートペアの二回以上の繰り返しがあった場合
		'\ud7c9[\udc00-\udfff]', // 特定のシリーズ除去
		'[0-9|*|#][\uFE0E-\uFE0F]\u20E3', // 数字系絵文字
		'[0-9|*|#]\u20E3', // 数字系絵文字
		'[©|®|\u2010-\u3fff][\uFE0E-\uFE0F]', // 環境依存文字や日本語との組み合わせによる絵文字
		'[\u2010-\u2FFF]', // 指や手、物など、単体で絵文字となるもの
		'\uA4B3' // 数学記号の環境依存文字の除去
	];

	// サロゲートペアの絵文字の文字コードを入れている
	const surrogatePairCode = [65038, 65039, 8205, 11093, 11035];

	// 文字コード配列を、|区切りで、文字列に変換＆正規表現を作成
	const reg = new RegExp(ranges.join('|'), 'g');

	// 貰った単語を、正規表現で検索して置換
	let retValue = in_value.replace(reg, '');

	// 一回の正規表現除去では除去しきないパターンがあるため、パターンにマッチする限り、除去を繰り返す
	while (retValue.match(reg)) {
		retValue = retValue.replace(reg, '');
	}

	// 二重で絵文字チェック（4バイト、サロゲートペアの残りカス除外）
	// retValueで二つの文字コードに分かれている場合、''で分割して、それぞれを処理している。
	retValue.split('').reduce((p, c) => {
		// 分割した部分の文字コードを取得
		const code = c.charCodeAt(0);

		// ？
		if (
			encodeURIComponent(c).replace(/%../g, 'x').length < 4 &&
			!surrogatePairCode.some((codeNum) => code == codeNum)
		) {
			return (p += c);
		} else {
			return p;
		}

		// 初期値として、''を定義している
	}, '');
}

//---------------------------------------------------------------------------------------------------------------------

function remove_emoji_test() {
	let worker_sheet = spreadSheet.getSheetByName('作業用');
	let emoji_array = ['😍', '👉'];
	let set_emoji_array = [['😍'], ['👉']];
	let sky_array = [];
	for (let i = 0; i < emoji_array.length; i++) {
		let sky = removeEmoji(emoji_array[i]);
		sky_array.push([sky]);
		console.log(emoji_array[i].length);
	}
	console.log(sky_array);
	worker_sheet.getRange(2, 2, sky_array.length, sky_array[0].length).setValues(sky_array);
	worker_sheet.getRange(2, 3, sky_array.length, sky_array[0].length).setValues(set_emoji_array);
}

//---------------------------------------------------------------------------------------------------------------------

// 1次元配列を二次元配列にする関数
function array1_to_array2_ver2(array) {
	let array2 = [];
	for (let i = 0; 0 < array.length; i) {
		array2.push(array.splice(i, 2));
	}
	return array2;
}

//---------------------------------------------------------------------------------------------------------------------

// typeと配列を受け取って、シートへ表示させる関数
function set_count_list_ver3(array, type) {
	// 配列が空の場合は何もせず、それ以外はカラムを指定
	if (array.length == 0) {
		return;
	} else {
		set_list_ver2(search_column_ver3(type));
	}

	// 繰り返し処理を避けるために、ブロック内に関数
	function set_list_ver2(column) {
		// 件数で降順に並び替える
		array = desc_sort_ver2(array);
		// シートに設置する
		output_sheet.getRange(3, column, array.length, 2).setValues(array);
	}
}

//---------------------------------------------------------------------------------------------------------------------

// array_2の件数を、二次元配列を、降順に並び替える関数
// https://tetsuooo.net/gas/2402/
function desc_sort_ver2(array) {
	function sort_by_count(a, b) {
		if (a[1] > b[1]) {
			return -1;
		} else if (a[1] < b[1]) {
			return 1;
		} else {
			return 0;
		}
	}

	array.sort(sort_by_count);

	return array;
}

//---------------------------------------------------------------------------------------------------------------------

// outputのpr列、afl列、org列、all列を探す関数
function search_column_ver3(title) {
	// outputシートのタイトルを取得
	let column_name_array = output_sheet.getRange(2, 1, 1, output_sheet.getLastColumn()).getValues()[0];

	// それぞれが何行目にあるか探している。indexOfは完全一致
	let pr_column = column_name_array.indexOf('PR_単語（＃）') + 1;
	let afl_column = column_name_array.indexOf('アフィ_単語（＃）') + 1;
	let org_column = column_name_array.indexOf('オーガニック_単語（＃）') + 1;
	let all_column = column_name_array.indexOf('全部_単語（＃）') + 1;
	let type_count_column = column_name_array.indexOf('各種の投稿（＃）') + 1;
	let kw_pr_column = column_name_array.indexOf('PR_単語（素）') + 1;
	let kw_afl_column = column_name_array.indexOf('アフィ_単語（素）') + 1;
	let kw_org_column = column_name_array.indexOf('オーガニック_単語（素）') + 1;
	let kw_all_column = column_name_array.indexOf('全部_単語（素）') + 1;

	// 見つからない場合、-1で帰ってくるけど、＋１しているので、０
	if (pr_column == 0 || afl_column == 0 || org_column == 0 || all_column == 0 || type_count_column == 0) {
		Browser.msgBox('列が見つからないよ！\n完全一致のため、タイトル行の名前変更があったのかも！');
	}

	// 引数のKWの列をreturnしている
	if (title == 'pr') {
		return pr_column;
	} else if (title == 'afl') {
		return afl_column;
	} else if (title == 'organic') {
		return org_column;
	} else if (title == 'all') {
		return all_column;
	} else if (title == 'type') {
		return type_count_column;
	} else if (title == 'kw_pr') {
		return kw_pr_column;
	} else if (title == 'kw_afl') {
		return kw_afl_column;
	} else if (title == 'kw_org') {
		return kw_org_column;
	} else if (title == 'kw_all') {
		return kw_all_column;
	}
}

//---------------------------------------------------------------------------------------------------------------------

//各種の投稿件数を円グラフで表示
function make_type_graph_ver1() {
	let type_column = search_column_ver3('type');
	// 全ての投稿数は割合に入れたくないので、３行目から
	let range = output_sheet.getRange(4, type_column, 3, 2);
	let pie_chart_builder = output_sheet
		.newChart()
		.addRange(range)
		.setChartType(Charts.ChartType.PIE)
		.setPosition(1, type_column + 3, 0, 0)
		.setOption('title', '各種の投稿件数の割合')
		.setOption('titleTextStyle', { color: 'black', bold: true })
		.build();
	// まだ図表が一つもない場合は、削除メソッドを実施しない
	if (output_sheet.getCharts().length !== 0) {
		remove_chart_ver1();
	}
	output_sheet.insertChart(pie_chart_builder);
}

//---------------------------------------------------------------------------------------------------------------------

function remove_chart_ver1() {
	let chart = output_sheet.getCharts()[0];
	console.log(chart);
	output_sheet.removeChart(chart);
}

//---------------------------------------------------------------------------------------------------------------------
