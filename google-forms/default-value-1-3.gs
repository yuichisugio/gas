/*
実現したいこと
１、受注報告WFで、同じ内容の入力を止めたい

具体的な方法、2パターン試す！
1､formで、自分が送信したform内容を複製できるURLを、その送信した人のGmailアドレスに送る仕組み
それで、複製が簡単になる。

2､Slackの受注報告の最後にも、複製用URLを載せても良いかも！
↓みたいな感じで！
<複製用URL>https://

  見本
  var form = FormApp.getActiveForm();
  var item = form.getItems()[i]
  var item_response = item.asTextItem().createResponse("hello");
  var create_response = form.createResponse();
  create_response.withItemResponse(item_response).toPrefilledUrl());
  // https://docs.google.com/forms/d/e/<formID>/viewform?usp=pp_url&entry.<ID>=hello
*/

//---------------------------------------------------------------------------------------------------------------------

// 回答データを取得して、複製URLを生成するバグがないバージョン！
function submitForm_ver3(e) {
	// フォームに設定しているItemを取得
	let items_form = test_form.getItems();

	// 回答データのItemを取得
	const items_response = e.response.getItemResponses();

	// フォームに事前に入れる内容を設定。
	// 日付は、↓の形式。日付は、dateオブジェクトである必要はあるので、
	// チェックリストは↓でさらに配列で囲む形式に変更する必要がある。
	// let word_array = ["俺！だぜ", "ジャスティス！", "中華ちまき", "はい", "2024-05-25", ["毎日", "週５"]];

	// 毎度for文内で宣言すると、入れる内容が上書きされてしまう。上書きされると貯められないので、外で宣言が必要！
	let response_form = test_form.createResponse();

	// 事前に入れる内容をfor文で作成
	for (let i = 0; i < items_response.length; i++) {
		// 先に、フォームに入れる
		let item_form;
		// 設問タイプを先に取得。それによって、asOOItem()を使い分ける
		let itemType = items_form[i].getType().toString();
		// 回答データを取得
		let answer = items_response[i].getResponse().toString();

		// 短文形式の設問
		if (itemType == 'TEXT') {
			item_form = items_form[i].asTextItem().createResponse(answer);

			// 複数行テキストタイプの設問
		} else if (itemType == 'PARAGRAPH_TEXT') {
			item_form = items_form[i].asParagraphTextItem().createResponse(answer);

			// ラジオボタンの設問
		} else if (itemType == 'MULTIPLE_CHOICE') {
			item_form = items_form[i].asMultipleChoiceItem().createResponse(answer);

			// 日付タイプの設問
		} else if (itemType == 'DATE') {
			// Dateオブジェクトで渡す必要があるため、変換している
			let responseDate = new Date(answer);
			item_form = items_form[i].asDateItem().createResponse(responseDate);

			// 複数選択チェックボックスの設問
		} else if (itemType == 'CHECKBOX') {
			console.log([answer]);
			// '毎日,週５'で渡されているので、['毎日','週５']に変換する
			let splitWordArray = answer.split(',');
			console.log(splitWordArray);
			item_form = items_form[i].asCheckboxItem().createResponse(splitWordArray);
		}

		// 毎回↓を行う
		response_form = response_form.withItemResponse(item_form);
	}

	// 最後に一回toPrefilledUrl()をして、入力用URLに変換している
	let outputUrl = response_form.toPrefilledUrl();

	// シートに表示！
	form_outputSheet.getRange(form_outputSheet.getLastRow() + 1, 1).setValue(outputUrl);
}

//---------------------------------------------------------------------------------------------------------------------

function submitForm_ver2(e) {
	// getItemResponsesは、回答データのアイテムを取得。getItemsは回答前の設定されている設問のアイテムを取得。クラスも違う
	const items_response = e.response.getItemResponses();
	let items_form = test_form.getItems();
	let prefilledUrl;
	let save_responses;

	for (let i = 0; i < items_response.length; i++) {
		let item = items_form[i];
		let answer = items_response[i].getResponse().toString();
		let preSetResponse;
		let itemType = item.getType().toString();

		if (itemType == 'TEXT') {
			preSetResponse = item.asTextItem().createResponse(answer);
		} else if (itemType == 'PARAGRAPH_TEXT') {
			preSetResponse = item.asParagraphTextItem().createResponse(answer);
		} else if (itemType == 'SCALE') {
			preSetResponse = item.asScaleItem().createResponse(answer);
		} else if (itemType == 'DATE') {
			preSetResponse = item.asDateItem().createResponse(answer);
		} else if (itemType == 'CHECKBOX') {
			preSetResponse = item.asCheckboxItem().createResponse(answer);
		}

		let create_response = test_form.createResponse();
		save_responses = create_response.withItemResponse(preSetResponse);
	}

	prefilledUrl = save_responses.toPrefilledUrl();
	form_outputSheet.getRange(form_outputSheet.getLastRow() + 1, 1).setValue(prefilledUrl);
}

//---------------------------------------------------------------------------------------------------------------------

// 俺が作ってみる！
function submitForm_ver1() {
	let prefillUrl = test_form.getPublishedUrl() + '?usp=pp_url&';

	//配列で得られる。response、itemResponsesで回答と質問の情報を取得
	const itemResponses = e.response.getItemResponses();

	// 質問と回答を取り出す
	for (let i = 0; i < itemResponses.length; i++) {
		const itemResponse = itemResponses[i];

		//これで「質問」が取れる
		const question = itemResponse.getItem().getTitle();

		//これで「回答」が取れる
		const answer = itemResponse.getResponse();

		// 設問IDを取得
		const id = itemResponse.getItem().getId();

		const encode_answer = encodeURIComponent(answer);

		// urlを作成
		prefillUrl = prefillUrl + 'entry.' + id + '=' + encode_answer + '&';
	}

	// 最後の & を削除
	prefillUrl = prefillUrl.slice(0, -1);

	// 出力シートにURLを書き込む
	form_outputSheet.getRange(form_outputSheet.getLastRow() + 1, 1).setValue(prefillUrl);
}

//---------------------------------------------------------------------------------------------------------------------

// デフォ表示URLを生成してくれるGAS
function generatePrefillUrls_ver1() {
	// form_inputSheet　　→シート
	// form_outputSheet →シート
	// form　→Google form

	// 前回の回答を取得
	var lastRow = form_inputSheet.getLastRow();
	var responses = form_inputSheet.getRange(lastRow, 1, 1, form_inputSheet.getLastColumn()).getValues()[0];

	// フォームのアイテムを取得
	var items = test_form.getItems();
	var prefillUrl = test_form.getPublishedUrl() + '?';

	// 前回の回答を使ってプレフィルURLを生成
	items.forEach((item, index) => {
		var itemResponse = responses[index + 1]; // 回答列はフォーム項目に対応
		if (itemResponse) {
			var itemId = item.getId();
			var entryId = itemIdToEntryId(itemId);
			var response = encodeURIComponent(itemResponse);
			prefillUrl += 'entry.' + entryId + '=' + response + '&';
		}
	});

	// 最後の & を削除
	prefillUrl = prefillUrl.slice(0, -1);

	// 出力シートにURLを書き込む
	form_outputSheet.getRange(form_outputSheet.getLastRow() + 1, 1).setValue(prefillUrl);
}

//---------------------------------------------------------------------------------------------------------------------

// すごくシンプルで、指定のワードだけ設問の一つ目に入れる関数
function generatePrefillUrls_ver2() {
	// フォームに設定しているItemを取得
	let items = test_form.getItems();

	// フォームに事前に入れる内容を設定
	let word_array = ['俺！だぜ（キラ', 'ジャスティス！'];

	// 毎度for文内で宣言すると、入れる内容が上書きされてしまう。上書きされると貯められないので、外で宣言が必要！
	let response = test_form.createResponse();

	// 事前に入れる内容をfor文で作成
	for (let i = 0; i < word_array.length; i++) {
		// アイテム内に入れる内容を取得
		let itemResponse = items[i].asTextItem().createResponse(word_array[i]);

		// response=responseで、上書きではなく、前の内容を貯めている
		response = response.withItemResponse(itemResponse);
	}

	// 最後に一回toPrefilledUrl()は最後しか入らないURLになる?varにして、グローバル変数にしてみる
	let outputUrl = response.toPrefilledUrl();

	// シートに表示！
	form_outputSheet.getRange(form_outputSheet.getLastRow() + 1, 1).setValue(outputUrl);
}

//---------------------------------------------------------------------------------------------------------------------

// タイプが違う場合も入るようにする！バグなしでOK！
function generatePrefillUrls_ver3() {
	// フォームに設定しているItemを取得
	let items = test_form.getItems();

	// フォームに事前に入れる内容を設定。
	// 日付は、↓の形式。日付は、dateオブジェクトである必要はあるので、
	// チェックリストは↓でさらに配列で囲む形式に変更する必要がある。
	let word_array = ['俺！だぜ', 'ジャスティス！', '中華ちまき', 'はい', '2024-05-25', ['毎日', '週５']];

	// 毎度for文内で宣言すると、入れる内容が上書きされてしまう。上書きされると貯められないので、外で宣言が必要！
	let response = test_form.createResponse();

	// 事前に入れる内容をfor文で作成
	for (let i = 0; i < word_array.length; i++) {
		let itemResponse;
		let itemType = items[i].getType().toString();

		if (itemType == 'TEXT') {
			itemResponse = items[i].asTextItem().createResponse(word_array[i]);
			response = response.withItemResponse(itemResponse);
		} else if (itemType == 'PARAGRAPH_TEXT') {
			itemResponse = items[i].asParagraphTextItem().createResponse(word_array[i]);
			response = response.withItemResponse(itemResponse);
		} else if (itemType == 'MULTIPLE_CHOICE') {
			itemResponse = items[i].asMultipleChoiceItem().createResponse(word_array[i]);
			response = response.withItemResponse(itemResponse);
		} else if (itemType == 'DATE') {
			let responseDate = new Date(word_array[i]);
			itemResponse = items[i].asDateItem().createResponse(responseDate);
			response = response.withItemResponse(itemResponse);
		} else if (itemType == 'CHECKBOX') {
			itemResponse = items[i].asCheckboxItem().createResponse(word_array[i]);
			response = response.withItemResponse(itemResponse);
		}
	}

	// 最後に一回toPrefilledUrl()をして、入力用URLに変換している
	let outputUrl = response.toPrefilledUrl();

	// シートに表示！
	form_outputSheet.getRange(form_outputSheet.getLastRow() + 1, 1).setValue(outputUrl);
}

//---------------------------------------------------------------------------------------------------------------------

function submitForm(e) {
	// フォームの情報を受け取って、タイトルを取得
	const title = e.source.getTitle();
	let message = '【' + title + '】に回答がありました\n\n';

	//配列で得られる。response、itemResponsesで回答と質問の情報を取得
	const itemResponses = e.response.getItemResponses();

	// 質問と回答を取り出す
	for (let i = 0; i < itemResponses.length; i++) {
		const itemResponse = itemResponses[i];

		//これで「質問」が取れる
		const question = itemResponse.getItem().getTitle();

		//これで「回答」が取れる
		const answer = itemResponse.getResponse();

		// 設問IDを取得
		const id = itemResponse.getItem().getId();

		//toString()で、(i+1)を文字列にしている
		//（今の）message = (前の)message + i+1 . question : answer
		message += (i + 1).toString() + '. ' + question + ':  ' + answer + '\n\n';
	}

	// slackに通知する
	postToSlack(message);
}

function postToSlack(message) {
	const url = 'https://hooks.slack.com/services/XXXXXXXXX/XXXXXXXXX/XXXXXXXXXXXXXXX';
	const payload = {
		text: message
	};

	const options = {
		method: 'post',
		contentType: 'application/json',
		payload: JSON.stringify(payload)
	};
	UrlFetchApp.fetch(url, options);
}

//---------------------------------------------------------------------------------------------------------------------

// 俺が作ってみる！
function submitForm_(e) {
	let prefillUrl = form.getPublishedUrl() + '?';
	let response = encodeURIComponent(itemResponse);
	prefillUrl += 'entry.' + entryId + '=' + response + '&';

	//配列で得られる。response、itemResponsesで回答と質問の情報を取得
	const itemResponses = e.response.getItemResponses();

	// 質問と回答を取り出す
	for (let i = 0; i < itemResponses.length; i++) {
		const itemResponse = itemResponses[i];

		//これで「質問」が取れる
		const question = itemResponse.getItem().getTitle();

		//これで「回答」が取れる
		const answer = itemResponse.getResponse();

		// 設問IDを取得
		const id = itemResponse.getItem().getId();

		const encode_answer = encodeURIComponent(answer);

		// urlを作成
		prefillUrl = prefillUrl + 'entry.' + id + '=' + encode_answer + '&';
	}

	// 最後の & を削除
	prefillUrl = prefillUrl.slice(0, -1);

	// 出力シートにURLを書き込む
	form_outputSheet.getRange(form_outputSheet.getLastRow() + 1, 1).setValue(prefillUrl);
}

//---------------------------------------------------------------------------------------------------------------------

// デフォ表示URLを生成してくれるGAS
function generatePrefillUrls_ver1() {
	// form_inputSheet　　→シート
	// form_outputSheet →シート
	// form　→Google form

	// 前回の回答を取得
	var lastRow = form_inputSheet.getLastRow();
	var responses = form_inputSheet.getRange(lastRow, 1, 1, form_inputSheet.getLastColumn()).getValues()[0];

	// フォームのアイテムを取得
	var items = form.getItems();
	var prefillUrl = form.getPublishedUrl() + '?';

	// 前回の回答を使ってプレフィルURLを生成
	items.forEach((item, index) => {
		var itemResponse = responses[index + 1]; // 回答列はフォーム項目に対応
		if (itemResponse) {
			var itemId = item.getId();
			var entryId = itemIdToEntryId(itemId);
			var response = encodeURIComponent(itemResponse);
			prefillUrl += 'entry.' + entryId + '=' + response + '&';
		}
	});

	// 最後の & を削除
	prefillUrl = prefillUrl.slice(0, -1);

	// 出力シートにURLを書き込む
	form_outputSheet.getRange(form_outputSheet.getLastRow() + 1, 1).setValue(prefillUrl);
}

//---------------------------------------------------------------------------------------------------------------------

// フォーム項目IDをエントリIDにマッピングする関数
function itemIdToEntryId(itemId) {
	var entryMapping = {
		FORM_ITEM_ID_1: 'ENTRY_ID_1',
		FORM_ITEM_ID_2: 'ENTRY_ID_2'
		// 必要に応じてフォーム項目IDとエントリIDのマッピングを追加
	};
	return entryMapping[itemId] || itemId;
}

//---------------------------------------------------------------------------------------------------------------------
