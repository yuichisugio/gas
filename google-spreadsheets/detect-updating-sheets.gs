// Google Apps Script code to check for updated cells in a specific Sheet tab daily at 09:00 and post a Slack notification
const SLACK_WEBHOOK_URL = 'https://hooks.slack.com/services/XXXXXXX/XXXXXXX/xxxxxxxxxxxxx';
const SPREADSHEET_ID = 'xxxxxxxxxxx-xxxxxxxxxxx'; // e.g., '1AbCdEfGh123...'
const SHEET_NAME = 'テンプレ';
const PROPERTIES_KEY = 'LAST_KNOWN_DATA';

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

function checkForUpdates() {
	const props = PropertiesService.getScriptProperties();
	const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
	const sheet = ss.getSheetByName(SHEET_NAME);
	if (!sheet) {
		Logger.log(`シート「${SHEET_NAME}」が見つかりません。`);
		return;
	}

	const dataRange = sheet.getDataRange();
	const values = dataRange.getValues();

	// 前回実行時に保存したデータを取得
	const lastKnownJson = props.getProperty(PROPERTIES_KEY);
	let lastKnownData = [];
	if (lastKnownJson) {
		lastKnownData = JSON.parse(lastKnownJson);
	}

	// 変更箇所(行, 列, 旧値→新値)をメモするための配列
	let changedCells = [];
	let updated = false;

	// (A) シートの行数・列数が変化しているかチェック
	const rowCountChanged = lastKnownData.length !== values.length;
	const colCountChanged = lastKnownData[0] && lastKnownData[0].length !== values[0].length;

	if (rowCountChanged || colCountChanged) {
		updated = true;
		changedCells.push(
			`Row/Columnの数が更新されました！\n` +
				`前回：行数${lastKnownData.length}で、列数が${(lastKnownData[0] || []).length}\n` +
				`現在：行数${values.length}で、列数が${values[0].length}`
		);
	}

	// (B) 行数・列数が同じならセルの内容を1つ1つ比較 (日付時差考慮)
	if (!rowCountChanged && !colCountChanged) {
		for (let row = 0; row < values.length; row++) {
			for (let col = 0; col < values[0].length; col++) {
				const currentValue = values[row][col];
				// 旧データに該当セルがなければ undefined とみなす
				const oldValue =
					lastKnownData[row] && typeof lastKnownData[row][col] !== 'undefined' ? lastKnownData[row][col] : undefined;

				// 日本時間に揃えた文字列として比較
				const currentNormalized = normalizeToJST(currentValue);
				const oldNormalized = normalizeToJST(oldValue);

				if (currentNormalized !== oldNormalized) {
					updated = true;
					let sendText = getLineDiffForBrSegments(oldNormalized, currentNormalized);
					for (let c = 0; c < sendText.length; c++) {
						changedCells.push(
							`${row + 1}行目・${col + 1}列目の更新内容\nAction：${sendText[c].action}\n変更内容：${
								sendText[c].textContent
							}\n------------------------------------------------------------------`
						);
					}
				}
			}
		}
	}

	// (C) 更新があれば Slack 通知
	if (updated) {
		const MAX_ITEMS_TO_SHOW = 300;
		let changesToShow = changedCells.slice(0, MAX_ITEMS_TO_SHOW);
		if (changedCells.length > MAX_ITEMS_TO_SHOW) {
			changesToShow.push(`...and ${changedCells.length - MAX_ITEMS_TO_SHOW} more changes`);
		}

		const text = [
			`<@UXXXXXXXXXXX> \n「${SHEET_NAME}」シートの変更通知だよ！`,
			'------------------------------------------------------------------',
			...changesToShow
		].join('\n');

		sendSlackNotification(text);
	} else {
		Logger.log(`更新は無かった！`);
	}

	// (D) 今回のシートデータを保存
	props.setProperty(PROPERTIES_KEY, JSON.stringify(values));
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

/**
 * Slack通知を送る
 */
function sendSlackNotification(message) {
	const payload = { text: message };
	const params = {
		method: 'post',
		contentType: 'application/json',
		payload: JSON.stringify(payload)
	};
	UrlFetchApp.fetch(SLACK_WEBHOOK_URL, params);
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

/**
 * 値を日本時間(JST)に揃えて文字列化する
 * - Date型なら Utilities.formatDate() で JST に変換した文字列を返す
 * - 日付文字列ならパースして JST に変換
 * - それ以外ならそのまま文字列として返す
 *
 * @param {*} value
 * @return {string}
 */
function normalizeToJST(value) {
	// (1) Date オブジェクトであれば JST 文字列に変換
	if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
		return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
	}

	// (2) 文字列であれば Date.parse() して有効なら JST 文字列に変換
	if (typeof value === 'string') {
		let parsed = Date.parse(value);
		if (!isNaN(parsed)) {
			let dateObj = new Date(parsed);
			return Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
		}
	}

	// (3) それ以外はそのまま文字列として返す
	return typeof value === 'undefined' ? '' : String(value);
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

/**
 * 旧HTMLと新HTMLの差分(行単位)を抽出して返すメイン関数。
 * - <br> 区切りで行配列を作成
 * - LCS(最長共通部分列)アルゴリズムで行ごとの差分を抽出
 * - 差分行のみを解析し、属性とテキストをまとめて返す
 *
 * @param {string} oldHtml 旧HTML文字列
 * @param {string} newHtml 新HTML文字列
 * @return {Array<Object>} 差分行を示すオブジェクトの配列
 *   例) [
 *      { action: 'removed', oldIndex: 2, newIndex: null,
 *        attributes: [...], textContent: "〇〇〇" },
 *      { action: 'added',   oldIndex: null, newIndex: 3,
 *        attributes: [...], textContent: "△△△" },
 *      ...
 *   ]
 */
function getLineDiffForBrSegments(oldHtml, newHtml) {
	// 1) <br>で分割して各行を配列に
	let oldSegments = splitByBr(oldHtml);
	let newSegments = splitByBr(newHtml);

	// 2) LCS(最長共通部分列)を求めるための DPテーブルを作成
	let dp = buildLCSMatrix(oldSegments, newSegments);

	// 3) DPテーブルから「どの行が削除され、どの行が追加されたか」差分リストを再構築
	let diffList = buildDiffList(oldSegments, newSegments, dp);

	// 4) 差分行だけ HTML解析（属性/テキスト抽出）して結果を返す
	let results = diffList.map(function (d) {
		let seg = d.action === 'removed' ? oldSegments[d.oldIndex] : newSegments[d.newIndex];

		let parsed = parseHtmlSegment(seg);

		// 差分情報にパース結果を付加
		return {
			action: d.action, // 'removed' or 'added'
			oldIndex: d.oldIndex, // 削除行の場合は行番号
			newIndex: d.newIndex, // 追加行の場合は行番号
			attributes: parsed.attributes,
			textContent: parsed.textContent
		};
	});

	return results;
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

/**
 * <br>, <br/>, <br /> を改行扱いにして分割し、配列を返す
 */
function splitByBr(html) {
	if (!html) return [];
	return html.split('\n');
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

/**
 * 2つの配列(oldArr, newArr)の最長共通部分列(LCS)の、長さを求めるためのDPテーブルを作成する
 * dp[i][j] は oldArr[0..i-1], newArr[0..j-1] のLCS長
 */
function buildLCSMatrix(oldArr, newArr) {
	let n = oldArr.length;
	let m = newArr.length;

	// (n+1) x (m+1) の2次元配列を0で初期化
	let dp = new Array(n + 1);
	for (let i = 0; i <= n; i++) {
		dp[i] = new Array(m + 1).fill(0);
	}

	for (let i = 1; i <= n; i++) {
		for (let j = 1; j <= m; j++) {
			if (oldArr[i - 1] === newArr[j - 1]) {
				dp[i][j] = dp[i - 1][j - 1] + 1;
			} else {
				dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
			}
		}
	}

	return dp;
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

/**
 * LCSのDPテーブルをもとに、どの行が削除され、どの行が追加されたか差分リストを復元する。
 *
 * 戻り値: 差分の配列
 *  [
 *    { action: 'removed', oldIndex: X, newIndex: null },
 *    { action: 'added',   oldIndex: null, newIndex: Y },
 *    ...
 *  ]
 * 同じ行(一致)は差分リストには含めない。
 */
function buildDiffList(oldArr, newArr, dp) {
	let i = oldArr.length;
	let j = newArr.length;
	let diff = [];

	// dp[n][m] から 0,0 まで辿りつつ差分を復元
	while (i > 0 && j > 0) {
		if (oldArr[i - 1] === newArr[j - 1]) {
			// 行が一致: 何も差分記録せず、両方進める
			i--;
			j--;
		} else if (dp[i - 1][j] > dp[i][j - 1]) {
			// old だけを1行戻したほうがLCSが長い -> old[i-1] は削除行
			diff.push({ action: 'removed', oldIndex: i - 1, newIndex: null });
			i--;
		} else {
			// new だけを1行戻したほうがLCSが長い -> new[j-1] は追加行
			diff.push({ action: 'added', oldIndex: null, newIndex: j - 1 });
			j--;
		}
	}

	// まだ oldArr が残っていれば削除された行
	while (i > 0) {
		i--;
		diff.push({ action: 'removed', oldIndex: i, newIndex: null });
	}

	// まだ newArr が残っていれば追加された行
	while (j > 0) {
		j--;
		diff.push({ action: 'added', oldIndex: null, newIndex: j });
	}

	// 逆順に溜まっているので反転
	return diff.reverse();
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

/**
 * 1つの行(HTML片)から「全タグの属性」と「テキスト本体」を抽出
 */
function parseHtmlSegment(segment) {
	let attributes = [];

	// タグを正規表現でスキャンし、属性を取り出す (<div style="..." id="..."> など)
	let tagRegex = /<([a-zA-Z][a-zA-Z0-9]*)([^>]*)>/g;
	let attrRegex = /([a-zA-Z_:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

	let matchTag;
	while ((matchTag = tagRegex.exec(segment)) !== null) {
		let attrString = matchTag[2];
		let matchAttr;
		while ((matchAttr = attrRegex.exec(attrString)) !== null) {
			let attrName = matchAttr[1];
			let attrValue = matchAttr[2] || matchAttr[3];
			attributes.push(attrName + '="' + attrValue + '"');
		}
	}

	// テキストコンテンツのみ取得 (タグ除去 & HTMLエンティティ変換)
	let textContent = segment
		.replace(/<[^>]+>/g, '')
		.replace(/&nbsp;/g, ' ')
		.replace(/&quot;/g, '"')
		.replace(/&amp;/g, '&')
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&apos;/g, "'")
		.trim();

	return { attributes: attributes, textContent: textContent };
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー
