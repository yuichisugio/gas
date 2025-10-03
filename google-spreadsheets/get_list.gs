// 来月の薬事チェック一覧を送る関数
function getPharmaceuticalCheckList() {
	// todayは今日の日付
	const today = new Date();
	const nextMonthDay = new Date(today.getFullYear(), today.getMonth() + 1, 1);
	// 次の月が何月が出力
	console.log('nextMonthDay: ' + Utilities.formatDate(nextMonthDay, Session.getScriptTimeZone(), 'yyyy/MM/dd'));

	// リマインドする際に送る情報を取得。
	const campaignId = new TeamXxxxArray('ID').array;
	const companyName = new TeamXxxxArray('会社名').array;
	const products = new TeamXxxxArray('商品名').array;
	const judgeNumber = new TeamXxxxArray('当選数').array;
	const status = new TeamXxxxArray('状況').array;
	const draftDeadline = new TeamXxxxArray('下書き〆').array;
	const postCheckDeadLine = new TeamXxxxArray('投稿チェック').array;
	const postedLegalDeadline = new TeamXxxxArray('法務チェック〆').array;
	const ccLegalDeadline = new TeamXxxxArray('原稿提出').array;
	const reviewRule = new TeamXxxxArray('チェック種別').array;

	// 案件数カウント
	let campaignCount = 0;
	let sumNumber = 0;

	// 送るテキスト
	let sendText = '';

	// 未作成の時に未作成を返す関数
	function getNewCampaignId(cpId) {
		let newCpId = '';
		if (cpId === 'CP') {
			newCpId = '未作成';
		} else {
			newCpId = cpId;
		}
		return newCpId;
	}

	// CC薬事チェックあり案件を探して、テキストに入れる
	for (let i = 0; i < campaignId.length; i++) {
		// TEAM_XXXX一覧で取得した日付をDate型に変更
		let newCcLegalDeadline = new Date(ccLegalDeadline[i]);
		let newDraftDeadline = new Date(draftDeadline[i]);
		let newPostCheckDeadline = new Date(postCheckDeadLine[i]);
		let newPostedLegalDeadline = new Date(postedLegalDeadline[i]);

		//「ステータスが未受注orレポート済み」or「チェック種別がーorCLor空欄」or「下書き〆切とCC法務〆切と事後投稿チェック〆切と事後CC法務〆切が日付ではない場合」はスキップ
		// isNaN(newCcLegalDeadLine.getTime())では、日付では無い場合はtrueになる
		if (
			status[i] == '未受注' ||
			status[i] == 'report済' ||
			status[i] == '失注' ||
			status[i] == '完了' ||
			reviewRule[i] == 'ー' ||
			reviewRule[i] == '' ||
			reviewRule[i] == 'CL' ||
			(isNaN(newCcLegalDeadline.getTime()) &&
				isNaN(newDraftDeadline.getTime()) &&
				isNaN(newPostCheckDeadline.getTime()) &&
				isNaN(newPostedLegalDeadline.getTime()))
		) {
			continue;
		}

		// 当選者数
		let shippingNumber = judgeNumber[i];

		// 審査の合計件数をカウントする関数
		function countLegal() {
			// 数字の文字列の場合のみ数値型に直すコード
			if (/^[+-]?\d+(\.\d+)?$/.test(shippingNumber)) {
				// 正規表現で数字のみの文字列かどうかを確認
				shippingNumber = Number(shippingNumber); // 数値型に変換
			} else {
				//改行されていることが多いので、replaceを使うためにString(judgeNumber[i])で文字列に変換している
				shippingNumber = String(judgeNumber[i]).replace(/\n/g, ' ');
			}

			// 数値型の場合のみ審査件数にカウント
			if (!isNaN(shippingNumber) && typeof shippingNumber === 'number' && isFinite(shippingNumber)) {
				sumNumber += shippingNumber;
			}
		}

		// 事後薬事の場合
		if (
			(reviewRule[i] == '事後CC' || reviewRule[i] == '事後両者') &&
			(newPostCheckDeadline.getMonth() == nextMonthDay.getMonth() ||
				newPostedLegalDeadline.getMonth() == nextMonthDay.getMonth())
		) {
			// 下書き〆切・LS法務〆切の日付形式を整える
			let formattedPostCheckDeadline = Utilities.formatDate(newPostCheckDeadline, 'JST', 'MM/dd');
			let formaattedPostedLegalDeadline = Utilities.formatDate(newPostedLegalDeadline, 'JST', 'MM/dd');

			// 案件数カウントを加算している
			campaignCount++;

			// ↑の条件に当てはまる時だけ関数を呼び出してカウントしている
			countLegal();

			const cpId = getNewCampaignId(campaignId[i]);

			// 送る文章を作成
			sendText +=
				'\n\n*' +
				companyName[i].replace(/\n/g, ' ') +
				'*' +
				'\n商品名：' +
				products[i].replace(/\n/g, ' ') +
				'\nキャンペーンID：' +
				cpId +
				'\n人数：' +
				shippingNumber +
				'\n審査期間：' +
				formattedPostCheckDeadline +
				'〜' +
				formaattedPostedLegalDeadline +
				'\n<https://slack.com/archives/xxxx/xxxx|※審査結果をメモに記載する案件>';

			// 「原稿提出\nLS法務〆」が10月 or 「下書き〆」が10月
		} else if (
			newCcLegalDeadline.getMonth() == nextMonthDay.getMonth() ||
			newDraftDeadline.getMonth() == nextMonthDay.getMonth()
		) {
			// 下書き〆切・LS法務〆切の日付形式を整える
			let formattedDraftDeacLine = Utilities.formatDate(newDraftDeadline, 'JST', 'MM/dd');
			let formaattedCcLegalDeadLine = Utilities.formatDate(newCcLegalDeadline, 'JST', 'MM/dd');

			// 案件数カウントを加算している
			campaignCount++;

			// ↑の条件に当てはまる時だけ関数を呼び出してカウントしている
			countLegal();

			const cpId = getNewCampaignId(campaignId[i]);

			// 送る文章を作成
			sendText +=
				'\n\n*' +
				companyName[i].replace(/\n/g, ' ') +
				'*' +
				'\n商品名：' +
				products[i].replace(/\n/g, ' ') +
				'\nキャンペーンID：' +
				cpId +
				'\n人数：' +
				shippingNumber +
				'\n審査期間：' +
				formattedDraftDeacLine +
				'〜' +
				formaattedCcLegalDeadLine;
		}
	}

	// 送る文章の最初に
	sendText =
		`<@XXXXXXX>\n *${
			nextMonthDay.getMonth() + 1
		}月の薬事チェック案件リスト作成してみたよ!*\n\n${campaignCount}案件あり、合計${sumNumber}名前後の審査になりそうです。` +
		sendText;

	// チャンネルに送信
	sendSlack(sendText, 'organic');
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// TEAM_XXXX一覧のヘッダー名を指定して、その列を１次元配列として取得する
class TeamXxxxArray {
	constructor(columnName) {
		// TEAM_XXXX一覧のデータのある一番下の行を取得。リマインドが必要な案件の開始行を指定。必要な範囲を計算
		this.startRow = 2450;
		this.searchRow = campaignSheet.getLastRow() - this.startRow;
		this.array = campaignSheet
			.getRange(this.startRow, getHeaders(campaignSheet, 2)[columnName], this.searchRow)
			.getValues()
			.flat();
	}
}
