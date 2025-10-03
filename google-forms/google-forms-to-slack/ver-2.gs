/**
 * 設定の定数を管理
 */
const FIRST_INCOMING_HOOK_URL = 'https://hooks.slack.com/services/xxxxxxxxx/xxxxxxxxx/xxxxxxxxx';

const FIRST_HEADER_MENTION = '<!subteam^xxxxxxxxx>';

const SECOND_INCOMING_HOOK_URL = 'https://hooks.slack.com/services/xxxxxxxxx/xxxxxxxxx/xxxxxxxxx';

const SECOND_HEADER_MENTION = '<!subteam^xxxxxxxxx> <!subteam^xxxxxxxxx> <!subteam^xxxxxxxxx>';

const THIRD_MENTION = '<@xxxxxxxxx> <@xxxxxxxxx> <@xxxxxxxxx> <@xxxxxxxxx>';

const FOURTH_MENTION = '<@xxxxxxxxx> <@xxxxxxxxx> <@xxxxxxxxx> <!subteam^xxxxxxxxx> <@xxxxxxxxx>';

const LIST_HEADER_MENTION = '<!subteam^xxxxxxxxx>';

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

/**
 * フォームの回答内容を解析してSlackに通知を送信する
 * @param {Object} e - formの入力内容が入る。
 * Google form送信がトリガー
 */
function autoSlack(e) {
	//フォームのデータを取得する。すべての質問と回答を取得する
	let itemResponses = e.response.getItemResponses();

	//個々の質問と回答を格納するための空配列を宣言する
	let questionAndAnswers = [];

	/**
	 * 1,2,3チャンネルへ飛ばすフラグ
	 */
	let workflowFlag = 0;

	//for文(ループ)で変数itemResponsesから個々の質問と回答を取得する
	for (let i = 0; i < itemResponses.length; i++) {
		//質問の文章を取得する
		let questionTitle = itemResponses[i].getItem().getTitle();

		//回答を取得する
		let answer = itemResponses[i].getResponse();

		//未回答の場合
		if (!answer) {
			// 回答に「未回答」を入れる
			questionAndAnswers.push('*≪' + questionTitle + '≫*\n 未回答\n');

			// 回答がある場合
		} else {
			// 「媒体」の回答のフォーム表示は不要のためスキップ。その代わり、フラグを立てる
			if (!questionTitle.includes('媒体')) {
				// 質問と回答を入れる
				questionAndAnswers.push('*≪' + questionTitle + '≫*\n' + answer + '\n');
			} else {
				// 1チャンネルの場合
				if (answer.includes('チャンネル-1')) {
					workflowFlag = 1;
				}

				// 2チャンネルの場合
				if (answer.includes('チャンネル-2')) {
					workflowFlag = 2;
				}
			}

			//XXがある場合、メンションを飛ばしている。
			if (answer.includes('XX')) {
				questionAndAnswers.push(THIRD_MENTION + '\n');
			}

			// XXがある場合、メンションを飛ばしている。
			if (answer.includes('XX')) {
				questionAndAnswers.push(FOURTH_MENTION + '\n');
			}
		}
	}

	const { url, header, username, icon } = getSendText(workflowFlag);

	//一次元配列questionAndAnswersに対してjoinメソッドを使って配列から文字列に変更する。区切り文字は改行"\n
	let body = header + questionAndAnswers.join('\n');

	//Slackに送信する
	sendSlack(url, body, username, icon);
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

/**
 * フラグごとに文言の差し替え
 * @param {number} flag - どの種類のテキストにするか返すためのフラグ
 */
function getSendText(flag) {
	if (flag === 0) {
		return {
			url: SECOND_INCOMING_HOOK_URL,
			header: '\n' + SECOND_HEADER_MENTION + '\n\n*title-1*\n\n',
			username: 'username-1',
			icon: ':icon-1:'
		};
	} else if (flag === 1) {
		return {
			url: SECOND_INCOMING_HOOK_URL,
			header: '\n' + LIST_HEADER_MENTION + '\n\n*title-2*\n\n',
			username: 'username-2',
			icon: ':icon-2:'
		};
	} else if (flag === 2) {
		return {
			url: FIRST_INCOMING_HOOK_URL,
			header: '\n' + FIRST_HEADER_MENTION + '\n\n*title-3*\n\n',
			username: 'username-3',
			icon: ':icon-3:'
		};
	}
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

/**
 * Slackに送信する関数
 */
function sendSlack(url, body, username, icon) {
	console.log(username);
	console.log(icon);

	// botに送る名前とアイコンと文章を格納
	let data = {
		username: username,
		icon_emoji: icon,
		text: body
	};

	// ↑の情報に追加で、Slackに送る方法や通信方法の指定を入れる
	let options = {
		method: 'post',
		contentType: 'application/json',
		payload: JSON.stringify(data)
	};

	// 指定URLに指定方法で送る
	UrlFetchApp.fetch(url, options);
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー
