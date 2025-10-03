// 例：回答数が100に達したらフォームを閉じるスクリプト
function checkResponseLimit() {
	let form = FormApp.getActiveForm(); // 現在のフォームを取得
	let responses = form.getResponses(); // 全ての回答を取得
	let responseCount = responses.length; // 回答数を取得
	let limit = 20; // ここに上限人数を設定

	if (responseCount >= limit) {
		form.setAcceptingResponses(false); // 回答の受付を停止
	}
}
