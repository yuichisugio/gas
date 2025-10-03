// スプシのヘッダー行の名前・順番が変わっている場合にアラートを出す関数
// リマインドは、ヘッダー名で列名を決めているので、列の並び替えはOKだけど、名前が変わるのはNGで気づきたいので、アラート関数を作成した！
// 名前が変わるだけだったら、別に、↑のリマインド関数は修正する必要がないので楽！↓のアラート関数を修正するだけでOK！
function detectChangeColumn() {
	const trueList = ['料金プラン', 'SNS', 'メニュー', 'オプション'];

	// 最終列を取得
	let lastColumn = campaignSheet.getLastColumn();
	// ヘッダー行をすべて（最終列まで）取得
	let headers = campaignSheet.getRange(2, 1, 1, lastColumn).getValues()[0];

	// 送る文章を作成
	let sendText = `<@UXXXXXXXX>\n*「一覧」の「キャンペーン一覧」タブのdetectChangeColumn()のお知らせ！*\n\n`;

	// 列名が変更されている列の数をカウント
	let changeCount = 0;

	// 以前の列名の数だけfor文を回して、違う列名があればsendTextに入れる
	for (let i = 0; i < trueList.length; i++) {
		if (trueList[i] != headers[i]) {
			// アラート文章
			sendText += `・ヘッダー行(2行目)の「${i + 1}列目」の「${trueList[i]}」列が編集されています！\n\n`;
			changeCount++;
		}
	}

	// 列名が変更されている列の数が1以上の場合はteamXxxxチャンネルに知らせる
	if (changeCount !== 0) {
		sendText += `※列の並び順 or セルの名前が変更されているかも！
    一覧が正しい場合は、<https://script.google.com/u/0/home/projects/XXXXXXXXXXXXXXXXXXXX|こちら>の「getTeamXxxxColumnList()」を使用して、この関数「detectChangeColumn()」の配列を正しい情報に更新してね！
    また、「リマインドGAS」や「スプシ→teamXxxxスケGAS」の列名も修正してね！`;
		// teamXxxxの個人チャンネルにアラートを出している
		sendSlack(sendText, 'teamXxxx');
		console.log(sendText);
	} else {
		sendText = `列名が変更された列はなかった！`;
		console.log(sendText);
	}
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// ↑の名前・列の並び順の変更を検知する関数の変更後の配列を返す関数
function getTeamXxxxColumnList() {
	// 最終列を取得
	let lastColumn = campaignSheet.getLastColumn();

	// ヘッダー行をすべて（最終列まで）取得
	let headers = campaignSheet.getRange(2, 1, 1, lastColumn).getValues()[0];

	// 配列を文字列に直した後も、↑の配列として貼り付けられるように、配列の形式で文字列化する
	let headersString = '["' + headers.join('","') + '"];';

	// 「\n」は文字列にしたり、スプシに書き出すと自動で改行されるため、改行されないように文字列で「\n」を入れている。「\n」だけだと自動で置き換えられるので、\を一つ増やしている
	headersString = headersString.replace(/\n/g, '\\n');

	// 「teamXxxx」の最終行の「商品名」列に、更新バージョンの配列を記載している。そこからコピペする。
	campaignSheet.getRange(campaignSheet.getLastRow(), 12, 1, 1).setValue(headersString);

	console.log(headers);
}
