// Google form送信がトリガーで、引数eにformの入力内容が入る。
function autoSlack(e) {

  //フォームのデータを取得する。すべての質問と回答を取得する
  let itemResponses = e.response.getItemResponses();

  //個々の質問と回答を格納するための空配列を宣言する
  let questionAndAnswers = [];

  //for文(ループ)で変数itemResponsesから個々の質問と回答を取得する
  for (let i = 0; i < itemResponses.length; i++) {

    //質問の文章を取得する
    let questionTitle = itemResponses[i].getItem().getTitle();

    //回答を取得する
    let answer = itemResponses[i].getResponse();

    //未回答の場合
    if (!answer) {
      questionAndAnswers.push("*≪" + questionTitle + "≫*\n 未回答\n");

    // 回答がある場合
    } else {
      questionAndAnswers.push("*≪" + questionTitle + "≫*\n" + answer + "\n");
    }

    //OOがある場合、別途メンションを飛ばしている。
    if (answer.includes("OO")) {
      questionAndAnswers.push("<@xxxxxxxxxxx> <@xxxxxxxxxxx> \n");
    }

    // yyが受注したら、別途メンションを飛ばしている。
    if (answer == "yy") {
      questionAndAnswers.push("<@xxxxxxxxxxx> \n");
    }
  }

  //Slackの宛先 (チャンネル)
  let url = "https://hooks.slack.com/services/xxxxxxxxx/xxxxxxxxx/xxxxxxxxx";

  // 冒頭のメンションを格納
  let metion = "<!subteam^xxxxxxxxx> <!subteam^xxxxxxxxx>";

  // ヘッダー文章
  let header = "\n" + metion + "\n\n*xxxxxxxxxの受注*\n\n"

  //一次元配列questionAndAnswersに対してjoinメソッドを使って配列から文字列に変更する。区切り文字は改行"\n
  let body = header + questionAndAnswers.join("\n");

  //Slackに送信する
  sendSlack(url, body);
}


// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー


// Slackに送信する関数
function sendSlack(url, body) {

  // botの表示名前
  const username = "xxxxxxxxx";

  // botのアイコン
  const icon = ":xxxxxxxxx:";

  // botに送る名前とアイコンと文章を格納
  let data = {
    "username": username,
    "icon_emoji": icon,
    "text": body
  }

  // ↑の情報に追加で、Slackに送る方法や通信方法の指定を入れる
  let options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(data)
  };

  // 指定URLに指定方法で送る
  UrlFetchApp.fetch(url, options);
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー
