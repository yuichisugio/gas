/*
改善点
１、すでに同じチャンネル名があった場合などのエラーハンドリング
２、グループメンションでも招待できる？
３、botが作成したチャンネルマネージャーは誰？
チャンネル作成後の設定する関数。チャンネル作成できるけどbotが作成者になるので、チャンネルマネージャーになるためにはSlack全体の管理者権限が必要になる

メモ
・基本botトークンを使用する。usertokenはユーザーとして振る舞いたい時のみ使用
・↓のページで必要なスコープを確認
https://api.slack.com/methods/chat.postMessage
*/

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// slackチャンネル作成に必要な関数をやり取りするTOP関数
function mainSlackChannel() {
	// エラーで中断させないようエラーハンドリング。
	try {
		// チャンネル作成　&　作成したチャンネルIDを取得
		// const channelId = createSlackChannel();
		// console.log("channelId: " + channelId);

		const channelId = slackSheet.getRange(6, 2).getValue();

		// メンバーを招待
		inviteMembersToChannel(channelId, token);

		// ブックマークを追加
		addBookmarkToChannel(channelId, token);

		// canvas作成
		createCanvas(channelId, token);

		// メッセージを送信
		const message = 'Canvas作成済み!';
		sendMessageToChannel(channelId, message, token);
	} catch (error) {
		console.log(`mainSlackChannel()で、エラーが発生しました: ${error.message}`);
		SpreadsheetApp.getUi().alert(`mainSlackChannel()で、エラーが発生しました: ${error.message}`);
	}
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 指定のslackチャンネルを作成する諸々の作業を行う関数
function createSlackChannel() {
	// シートからチャンネル名を取得
	let channelName = slackSheet.getRange(3, 2).getValue().trim();
	channelName = sanitizeTeamName(channelName);

	// チャンネル名が未記載の場合は、ダイアログを表示
	if (!channelName) {
		SpreadsheetApp.getUi().alert(
			'createSlackChannel()で、チャンネル名が指定されていません。A1セルにチャンネル名を入力してください。'
		);
		console.log('createSlackChannel()で、チャンネル名が指定されていません。A1セルにチャンネル名を入力してください。');
		return;
	}

	// slackチャンネル名の無効な文字をチェック
	const invalidChars = /[A-Z\s\.']/; // 大文字、スペース、ピリオド、アポストロフィを除外
	if (invalidChars.test(channelName)) {
		SpreadsheetApp.getUi().alert(
			'チャンネル名に無効な文字が含まれています。小文字の英数字、ハイフン、アンダースコアのみを使用してください。'
		);
		console.log(
			'createSlackChannel()で、チャンネル名に無効な文字が含まれています。小文字の英数字、ハイフン、アンダースコアのみを使用してください。'
		);
		return;
	}

	try {
		// チャンネルを作成
		let channelId = createSlackChannelAPI(channelName, token);
		return channelId;
	} catch (error) {
		SpreadsheetApp.getUi().alert(`createSlackChannel()で、エラーが発生しました: ${error.message}`);
		console.log(`createSlackChannel()で、エラーが発生しました: ${error.message}`);
		return;
	}
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// slackチャンネルの命名規則に合うように修正する
function sanitizeTeamName(teamName) {
	// チーム名を処理し、記号を削除/置換し、80文字以内に短縮
	let sanitized = '';
	// 大文字アルファベットは小文字に変換
	sanitized = teamName.toLowerCase();
	// 日本語（ひらがな、カタカナ、漢字）、長音符（ー）、アルファベット、ハイフン（-）、アンダースコア（_）以外の記号を削除
	sanitized = sanitized.replace(/[^ぁ-んァ-ン一-龠ーa-zA-Z0-9-_]/g, '');
	// 80文字以内に短縮。80文字目以降を削除
	sanitized = sanitized.substring(0, 80);
	return sanitized;
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// Slackチャンネルを作成する専用関数
function createSlackChannelAPI(channelName, token) {
	const url = 'https://slack.com/api/conversations.create';

	const payload = {
		name: channelName,
		is_private: false
	};

	const options = {
		method: 'post',
		contentType: 'application/json',
		headers: {
			Authorization: `Bearer ${token}`
		},
		payload: JSON.stringify(payload),
		// muteHttpExceptions は、エラー応答（HTTPステータスコードが 200 ではない応答）をどう処理するかに関するオプション。デフォルトでは、サーバーからのHTTPステータスコードが200番台以外（つまり、400や500などのエラーが発生している場合）のレスポンスが返ってきたときに、Google Apps Scriptは自動的に**例外（エラー）**をスローします。この場合、スクリプトはそこで停止し、通常のレスポンスデータは取得できなくなります。しかし、muteHttpExceptions: true を設定すると、エラーレスポンスが返ってきても、スクリプトが停止せずに続行します。このオプションを有効にすると、エラーレスポンスも通常のレスポンスと同様に取得することができます。ステータスコードに応じたカスタムメッセージをログに出力したい時など。404の時は「」とか
		muteHttpExceptions: true
	};

	const response = UrlFetchApp.fetch(url, options);
	const result = JSON.parse(response.getContentText());

	if (!result.ok) {
		console.log(`createSlackChannelAPI()で、チャンネルの作成に失敗しました: ${result.error}`);
		throw new Error(`createSlackChannelAPI()で、チャンネルの作成に失敗しました: ${result.error}`);
	}

	console.log('チャンネル作成成功！');
	return result.channel.id;
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// シートからメンバーIDを取得する関数。slack IDを「vwe,vweb,wege」カンマ区切りで返したい
function getInviteMembersIds() {
	// チャンネル招待するメンバー
	let inviteMembersIds = '';

	// ID取得の為の、メアド取得の為の、TEAM_XXXX一覧とメアドの突合リストを作成
	let flowMembersArray = [['member_xxxx', 'member_xxxx@example.com']];
	console.log('到達！');

	// pmメンバーを取得 & 招待メンバーに追加
	let channelRow = slackSheet.getRange(2, 2).getValue();
	console.log('channelRow: ' + channelRow);
	let pmMembers = campaignSheet.getRange(channelRow, 6).getValue();
	console.log('pmMember: ' + pmMembers);
	// 営業メンバーの取得　＆　招待メンバーに追加
	let salesMembers = campaignSheet.getRange(channelRow, 5).getValue();
	console.log('salesMember:' + salesMembers);
	// 追加メンバー用
	let additonalMembers = slackSheet.getRange(5, 2).getValue();
	for (let i = 0; i < flowMembersArray.length; i++) {
		console.log(flowMembersArray[i]);
		console.log(flowMembersArray[i][0]);
		console.log(flowMembersArray[i][1]);
		// PMメンバー or 営業メンバーがいるか判断
		if (
			pmMembers.includes(flowMembersArray[i][0]) ||
			salesMembers.includes(flowMembersArray[i][0]) ||
			additonalMembers.includes(flowMembersArray[i][0])
		) {
			// slackIDを取得
			let slackId = getUserIdByEmail(flowMembersArray[i][1]);
			// slackIDがあれば、IDsに追加
			if (slackId) {
				inviteMembersIds += ',' + slackId;
			}
		}
		console.log('到達！_2');
	}
	console.log('到達3');

	// チャンネル招待の固定メンバーを取得。↓は、OPE・KOLのグループメンション
	const staticMembers = 'SXXXXXXXXXXX,SXXXXXXXXXXX';
	// チャンネル招待するメンバー
	inviteMembersIds += ',' + staticMembers;
	// 最初or最後の「,」を抜く
	inviteMembersIds = inviteMembersIds.replace(/^,|,$/g, '');

	console.log(inviteMembersIds);
	console.log('getInviteMembersIds()で、招待メンバーリストの作成に成功しました！');
	return inviteMembersIds;
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

//メールアドレスで検索して、ユーザーIDを返す関数
function getUserIdByEmail(email) {
	//APIのURL
	let url = 'https://slack.com/api/users.lookupByEmail';

	let payload = {
		token: token, //botトークン
		email: email //検索したいメールアドレス
	};

	let options = {
		method: 'GET',
		payload: payload,
		headers: {
			contentType: 'application/x-www-form-urlencoded'
		}
	};

	let json_data = UrlFetchApp.fetch(url, options); //APIリクエスト実行と結果の格納
	json_data = JSON.parse(json_data.getContentText()); //結果はJSONデータで返されるのでデコード

	let user_id;

	if (json_data['ok']) {
		//boolean型でtrue or falseが格納されています
		user_id = json_data['user']['id']; //trueの場合返答されたデータからユーザーIDを抽出
	} else {
		user_id = false; //falseの場合null(文字列)を格納
	}

	//ユーザーID(or false)を返却
	return user_id;
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// メンバーをチャンネルに招待する関数
function inviteMembersToChannel(channelId, token) {
	// 招待するためのslackAPIのURL
	const url = 'https://slack.com/api/conversations.invite';

	// 招待リストを作成
	const inviteMembersIds = getInviteMembersIds();

	const payload = {
		channel: channelId,
		users: inviteMembersIds,
		force: true //有効なユーザーのみ招待するかどうか、falseなら有効ではない場合エラーになるっぽい
	};

	const options = {
		method: 'post',
		contentType: 'application/json',
		headers: {
			Authorization: `Bearer ${token}`
		},
		payload: JSON.stringify(payload),
		muteHttpExceptions: true
	};

	const response = UrlFetchApp.fetch(url, options);
	const result = JSON.parse(response.getContentText());

	if (!result.ok) {
		console.log(`inviteMembersToChannel()で、メンバーの招待に失敗しました: ${result.error}`);
		throw new Error(`inviteMembersToChannel()で、メンバーの招待に失敗しました: ${result.error}`);
	}
	console.log('inviteMembersToChannel()で、招待に成功！');
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// チャンネルにメッセージを送信する関数
function sendMessageToChannel(channelId, message, token) {
	// postのAPI URLを設定
	const url = 'https://slack.com/api/chat.postMessage';

	const payload = {
		channel: channelId,
		text: message
	};

	const options = {
		method: 'post',
		contentType: 'application/json',
		headers: {
			Authorization: `Bearer ${token}`
		},
		payload: JSON.stringify(payload),
		muteHttpExceptions: true
	};

	const response = UrlFetchApp.fetch(url, options);
	const result = JSON.parse(response.getContentText());

	if (!result.ok) {
		console.log(`sendMessageToChannel()で、メッセージの送信に失敗しました: ${result.error}`);
		throw new Error(`sendMessageToChannel()で、メッセージの送信に失敗しました: ${result.error}`);
	}
	console.log('sendMessageToChannel()で、メッセージの送信完了！');
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// チャンネルにブックマークを追加する関数
function addBookmarkToChannel(channelId, token) {
	// ブックマークを追加するAPI URL
	const url = 'https://slack.com/api/bookmarks.add';

	// ブックマークに入れる内容を取得
	let channelRow = slackSheet.getRange(2, 2).getValue();
	const driveUrl = campaignSheet.getRange(channelRow, 14).getValue().trim();
	const scheduleUrl = slackSheet.getRange(4, 2).getValue().trim();

	// ブックマークに入れる内容を全部入れる
	let bookmarkContents = [
		[driveUrl, 'クライアントフォルダ'],
		[scheduleUrl, '最新スケジュール']
	];

	// ブックマークに入れるコンテンツ分だけ回す。
	for (let i = 0; i < bookmarkContents.length; i++) {
		// URLがある場合は送る
		if (bookmarkContents[i][0]) {
			const payload = {
				channel_id: channelId,
				title: bookmarkContents[i][1],
				type: 'link',
				link: bookmarkContents[i][0]
			};

			const options = {
				method: 'post',
				contentType: 'application/json',
				headers: {
					Authorization: `Bearer ${token}`
				},
				payload: JSON.stringify(payload),
				muteHttpExceptions: true
			};

			const response = UrlFetchApp.fetch(url, options);
			const result = JSON.parse(response.getContentText());

			if (!result.ok) {
				console.log(`addBookmarkToChannel()で、ブックマークの追加に失敗しました: ${result.error}`);
				throw new Error(`addBookmarkToChannel()で、ブックマークの追加に失敗しました: ${result.error}`);
			}
			console.log('addBookmarkToChannel()で、ブックマークの追加は完了!');
		}
	}
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// Canvasを作成する関数
function createCanvas(channelId, token) {
	//  canvasを作成するAPI URL
	const url = 'https://slack.com/api/conversations.canvases.create';

	// let templateText = `## 確認したいこと\n* new\n* new\n\n## OPEタスク\n* new\n* new\n\n## KOLタスク\n* new\n* new\n\n## 確認済み\n* new\n* new\n\n## スケジュール条件\n* new\n* new\n\n## 募集ページ条件\n* new\n* new\n* やること\n* new\n* 事前アンケート\n * new`;
	let templateText = getTemplateText();

	const payload = {
		channel_id: channelId,
		document_content: {
			type: 'markdown',
			markdown: templateText
		}
	};

	const options = {
		method: 'post',
		contentType: 'application/json',
		headers: {
			Authorization: `Bearer ${token}`
		},
		payload: JSON.stringify(payload),
		muteHttpExceptions: true
	};

	const response = UrlFetchApp.fetch(url, options);
	const result = JSON.parse(response.getContentText());

	if (!result.ok) {
		console.log(`createCanvas()で、メッセージの送信に失敗しました: ${result.error}`);
		throw new Error(`createCanvas()で、メッセージの送信に失敗しました: ${result.error}`);
	}
	console.log('createCanvas()で、canvasの作成完了!');
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// canvasにテンプレで入れる文章を作成。
// markdown記法の改行コードなしで、バックティントにしないと入れ子が反映されなかった
// 改行コードは一文につき一つしか使用できなかった
function getTemplateText() {
	let templateText = `
  ## 確認したいこと
  * 
  * 
  `;

	return templateText;
}
