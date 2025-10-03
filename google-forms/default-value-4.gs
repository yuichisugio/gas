/*
実現したいこと
１、受注報告WFで、同じ内容の入力を止めたい

具体的な方法、2パターン試す！
1､formで、自分が送信したform内容を複製できるURLを、その送信した人のGmailアドレスに送る仕組み
それで、複製が簡単になる。

2､Slackの報告の最後にも、複製用URLを載せても良いかも！
↓みたいな感じで！
<複製用URL>https://

メモ
コンテナバインドではなくても、その関数をトリガーに設定するだけで、eに情報が入ってくるらしい！便利！
３、URLの文字数は、2000文字前後までしか無理なので、制限があるので注意！
*/


//---------------------------------------------------------------------------------------------------------------------


// Slack・送信者のGmailメアド・outputシートの三箇所に、複製用
function submitForm_ver4(e) {

  // 複製用URLを作成
  let prefillUrl = make_prefilledUrl_ver1(e);

  // スプシに表示させる内容を作成
  let sheetMessage = make_sheetBody_ver1(e, prefillUrl);

  // スプシに文章を記載
  postToSpreadSheet(sheetMessage);

  // Slackに送る受注報告WFの文章を作成
  let slackMessage = make_slackBody_ver1(e, prefillUrl);

  // Slackに送る依頼
  postToSlack(slackMessage);

  // Gmailに送る文章を作成
  let gmailMessage = make_gmailBody_ver1(e, prefillUrl);

  // Gmailに送る依頼
  postToGmail(gmailMessage);

}


//---------------------------------------------------------------------------------------------------------------------


function make_prefilledUrl_ver1(responses) {

  // フォームに設定しているItemを取得
  let items_form = test_form.getItems()

  // 回答データのItemを取得
  const items_response = responses.response.getItemResponses();

  // 毎度for文内で宣言すると、入れる内容が上書きされてしまう。上書きされると貯められないので、外で宣言が必要！
  let response_form = test_form.createResponse()

  // 事前に入れる内容をfor文で作成
  for (let i = 0; i < items_response.length; i++) {

    // 先に、フォームに入れる文章の変数を宣言
    let item_form;

    // 設問タイプを先に取得。それによって、asOOItem()を使い分ける
    let itemType = items_form[i].getType().toString();

    // 回答データを取得
    let answer = items_response[i].getResponse().toString();

    // 短文形式の設問
    if (itemType == "TEXT") {
      item_form = items_form[i].asTextItem().createResponse(answer);

      // 複数行テキストタイプの設問
    } else if (itemType == "PARAGRAPH_TEXT") {
      item_form = items_form[i].asParagraphTextItem().createResponse(answer);

      // ラジオボタンの設問
    } else if (itemType == "MULTIPLE_CHOICE") {
      item_form = items_form[i].asMultipleChoiceItem().createResponse(answer);

      // 日付タイプの設問
    } else if (itemType == "DATE") {
      // Dateオブジェクトで渡す必要があるため、変換している
      let responseDate = new Date(answer);
      item_form = items_form[i].asDateItem().createResponse(responseDate);

      // 複数選択チェックボックスの設問
    } else if (itemType == "CHECKBOX") {
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

  // 複製用URLを返す
  return outputUrl;
}


//---------------------------------------------------------------------------------------------------------------------


// Gmailに送る文章を作成
function make_gmailBody_ver1(responses, prefillUrl) {

  // 入力者のメアドを取得
  let emailAddress = responses.response.getRespondentEmail();

  // 件名
  let subject = '受注報告ワークフローの入力完了のお知らせ!';

  // メールの中身
  let body = `受注報告ワークフローの入力ありがとうございます!\n\n複製して、前回の内容が入った状態で再度入力したい場合は、↓のリンクを押してね!\n${prefillUrl}`

  // 記載内容を配列に入れる
  let gmail_message = [emailAddress, subject, body];

  // 配列で返す
  return gmail_message;
}


//---------------------------------------------------------------------------------------------------------------------


// 受け取った文章をGmailに送信
function postToGmail(message) {

  // 受け取った配列を、メアド・件名、文書を文字列として取得
  let emailAddress = message[0];
  let subject = message[1];
  let body = message[2];

  // 実際にGmailに送る
  GmailApp.sendEmail(emailAddress, subject, body);
}


//---------------------------------------------------------------------------------------------------------------------


// Slackに送る文章を作成
function make_slackBody_ver1(responses, prefillUrl) {

  // フォームの情報を受け取って、タイトルを取得
  const title = responses.source.getTitle();

  // 送信者にもメンション飛ばしたい！ので、送信者を取得
  let userEmail = responses.response.getRespondentEmail();

  //メアドからslack_User_IDを取得
  let user_id = findUserIdByEmail(userEmail);

  // 入力者のメンション変数を宣言
  let mention = "";

  //送信者のslackユーザーを招待、メンションの準備をする。居ない場合はメールアドレス
  if (user_id == null) {
    mention = userEmail

  } else {
    mention = " <@" + user_id + ">" //送信者へメンションする書式
  }

  // 受注報告WFのタイトルを作成
  const headerMessage = `<@UXXXXXXXXXXX>\n【${title}】に入力があったよ！\n\n*≪入力者名≫*\n${mention}\n\n`;

  //配列で得られる。response、itemResponsesで回答と質問の情報を取得
  const itemResponses = responses.response.getItemResponses();

  // 文章を入れる配列
  let questionAndAnswers = [];

  // 質問と回答を取り出す
  for (let i = 0; i < itemResponses.length; i++) {

    // 一つづつの質問と回答のセットを取り出す
    let itemResponse = itemResponses[i];

    //「質問」をとる
    let question = itemResponse.getItem().getTitle();

    //「回答」をとる
    let answer = itemResponse.getResponse();

    if (!answer) {
      questionAndAnswers.push("*≪" + question + "≫*\n 未回答");
    } else {
      questionAndAnswers.push("*≪" + question + "≫*\n" + answer);
    }
  }

  // 複製用URLを末尾に入れる
  let prefilledUrlMessage = `*≪複製用URL≫*\n複製して、前回の内容が入った状態で再度入力したい場合は、↓のリンクを押してね!\n${prefillUrl}`
  questionAndAnswers.push(prefilledUrlMessage);

  //Slackの本文。一次元配列questionAndAnswersに対してjoinメソッドを使って文字列を作成する。区切り文字は改行"\n"
  let slack_message = headerMessage + questionAndAnswers.join("\n\n");

  // slack文章を返す
  return slack_message;

}


//---------------------------------------------------------------------------------------------------------------------


//メールアドレスで検索して、ユーザーIDを返す関数
function findUserIdByEmail(email) {

  // botのtoken
  const botToken = "xoxb-xxxxxxxxxxxxx";

  //APIのURL
  let url = "https://slack.com/api/users.lookupByEmail";
  let payload = {
    "token": botToken, //botトークン
    "email": email //検索したいメールアドレス
  }

  let options = {
    "method": "GET",
    "payload": payload,
    "headers": {
      "contentType": "application/x-www-form-urlencoded"
    }
  }

  let json_data = UrlFetchApp.fetch(url, options); //APIリクエスト実行と結果の格納
  json_data = JSON.parse(json_data.getContentText()); //結果はJSONデータで返されるのでデコード

  let user_id;

  if (json_data["ok"]) { //boolean型でtrue or falseが格納されています
    user_id = json_data["user"]["id"]; //trueの場合返答されたデータからユーザーIDを抽出

  } else {
    user_id = null; //falseの場合null(文字列)を格納
  }

  //ユーザーID(or null)を返却
  return user_id 
}


//---------------------------------------------------------------------------------------------------------------------


function postToSlack(message) {

  // 受注報告を送るチャンネルのincoming-webhook URL
  const url = "https://hooks.slack.com/services/xxxxxxxxx/xxxxxxxxx/xxxxxxxxxxxxx";

  // ペイロードの設定
  const payload = {
    'text': message
  }

  // ペイロードも含めたオプションの設定
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  }

  // 送る
  UrlFetchApp.fetch(url, options);
}


//---------------------------------------------------------------------------------------------------------------------


// シートに記載する文章を作成する関数
function make_sheetBody_ver1(responses, prefillUrl) {
  
  // 送信時間を取得
  let postTimeStamp = responses.response.getTimestamp();

  // 送信者を取得
  let emailAddress = responses.response.getRespondentEmail();

  // 複製用URLの表示名
  let displayName = "複製用リンク";

  // ハイパーリンクを作成
  let hyperlinkFormula = '=HYPERLINK("' + prefillUrl + '", "' + displayName + '")';

  // ハイパーリンクを作る関数文字列、送信日時、送信者のメアドを配列に入れて返す
  return [hyperlinkFormula, postTimeStamp, emailAddress];
}


//---------------------------------------------------------------------------------------------------------------------


// 引数の文章を、シートに記載！
function postToSpreadSheet(message) {

  // 二次元配列にしている
  let body = [message];

  // 複製用ボタンURLのハイパーリンク・フォームの送信時間をシートに記載
  form_outputSheet.getRange(form_outputSheet.getLastRow() + 1, 1, body.length, body[0].length).setValues(body);
}
