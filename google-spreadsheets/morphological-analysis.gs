/* 
メモ
800件でAPIで怒られた
1分間で300件だとダメなので、1分以内で300件行きそうなら、sleep()する仕組みを作る？
2個を一個にしたら、文章が多いと怒られた
100〜200投稿なら可能っぽい 
*/


//---------------------------------------------------------------------------------------------------------------------


function kw_search_count_ver6() {

  // キャプションの列の最終行を取得している。getlastrow()だと、投稿SNS検知の「投稿なし」に反応して無駄に取得してしまう
  let caption_lastRow = get_designation_last_row(2)

  // 2行目、１列目、データがある行まで、３列目まで取得
  let kw_captions_array = input_sheet.getRange(2, 1, caption_lastRow, 3).getValues();

  // 件数が多い場合は、ダイアログを表示して止めている。returnをすればメイン関数も止まる。
  if (kw_captions_array.length >= 200) {
    Browser.msgBox(`現在inputに入れている件数が${kw_captions_array.length}件です！\\n件数が多すぎるため、150件以下にしてね！`);
    return;
  }

  // それぞれのカウント配列を作成
  let kw_pr_count_array = [];
  let kw_afl_count_array = [];
  let kw_org_count_array = [];
  let kw_all_count_array = [];

  // 追加のPR_KWを聞いている
  let dialog_pr_kw = get_aditional_pr_kw();

  // PR判定KW
  let pr_regexPattern_reg = get_regexPattern("pr", dialog_pr_kw);
  console.log(pr_regexPattern_reg);

  // アフィ判定KW
  let afl_regexPattern_reg = get_regexPattern("afl");
  console.log(afl_regexPattern_reg);

  // 1つの投稿づつYAHOO APIに投げている
  for (let j = 0; j < kw_captions_array.length; j++) {

    // キャプション
    let kw_caption = kw_captions_array[j][1];

    // 投稿SNS
    let sns = kw_captions_array[j][2];

    // 正規表現に当てはまるか確認。当てはまる場合は配列を返す
    // PR判定用
    let pr_regex_array = get_regexPattern_array(pr_regexPattern_reg, kw_caption);
    // アフィ判定
    let afl_regex_array = get_regexPattern_array(afl_regexPattern_reg, kw_caption);

    // PR判定・アフィ判定用のbooleanを作成＆判定
    let kw_pr_boolean = judge_pr_afl("pr", pr_regex_array, sns);
    let kw_afl_boolean = judge_pr_afl("afl", afl_regex_array, sns);

    // 確認済みの投稿件数を出している
    console.log("確認済みの投稿件数：" + j);

    // APIを叩いていて、tokensの中身だけ取得している
    let response_tokens_array = yahooApiRequest(kw_caption);

    // 返ってきた内容が配列ではない場合、エラーだと判断して、メッセージボックスを表示
    if (!Array.isArray(response_tokens_array)) {
      Browser.msgBox(response_tokens_array, Browser.Buttons.OK);
      console.log(response_tokens_array);
      return;
    }
    console.log(response_tokens_array);

    // response_tokens_arrayから、単語だけを取り出している
    for (let s = 0; s < response_tokens_array.length; s++) {

      let target_word = response_tokens_array[s][0];
      let target_hinshi = response_tokens_array[s][3];

      /* 全部がremoveされるので、絵文字検知したいけど、一旦Pending。未定義語全部が、非表示になっている
      if (target_hinshi == "未定義語") {
        target_word = removeEmoji(target_word);
      }*/

      // 単語でスキップするリスト。\nを入れることで空欄がなくなった！
      let skip_word_array = [
        "😊", "⋱", "⋰", "゚", "▶", "𓂃", "[", "]", "˗", "●", "✂", "⌒", "☆", "した", "よう", "🤍", "━", "🥰", "🙆", "🔥", "🪞",
        "😍", "𓏸", "の", "な", "さ", "こと", "̄", "̄", "して", "し", "この", "だ", "なる", "です", ">", "🥹", "♀", "⚠", "🫶", "🥺", "💕",
        "👍", "💄", "💎", "こちら", "する", "", "", '︎', '⁡', " ", "  ", "   ", "    ", null, undefined, '', '⠀', "⠀⠀",
        "'ᅠ'", "ᅠ", "ᅠ", "\n", "\r", "\s", "/", "\\", ".", "#", "@", "_", "-", "̅", "ー", "(", ")", "=", "*", "♡", "!", "?",
        ":", "→", "一", "✨", "┈", "•", "°", "☑", "⭐", "☺", "·", "✔", "◎", "🏻", "⚪", "❤", "♪", "✅"
      ];
      let skip_hinshi_array = ["助詞", "接尾辞", "副助詞", "特殊", "助動詞", "接頭辞", "接続詞"];

      // 入れる値が配列内の空欄の場合、スキップ
      if (skip_word_array.includes(target_word)) {
        continue;
      }
      if (skip_hinshi_array.includes(target_hinshi)) {
        continue;
      }
      if (target_word == null || target_word == undefined) {
        continue;
      }
      // 空欄削除作戦。何個スペース入れていても対応できるバージョン！
      if (target_word.charAt(0) == "" && target_word.charAt(1) == "") {
        continue;
      }

      // それぞれが、cout_arrayにあるかどうか見て、ある場合は、そのインデックス番号、ない場合は-1を返す
      let kw_afl_number = kw_afl_count_array.indexOf(target_word);
      let kw_pr_number = kw_pr_count_array.indexOf(target_word);
      let kw_org_number = kw_org_count_array.indexOf(target_word);
      let kw_all_number = kw_all_count_array.indexOf(target_word);

      // PR・AFL・ORG判定した投稿を、それそれの種類でカウント
      if (kw_afl_boolean) {
        if (kw_afl_number >= 0) {
          kw_afl_count_array[kw_afl_number + 1]++;
        } else {
          kw_afl_count_array.push(target_word, 1);
        }
      } else if (kw_pr_boolean) {
        if (kw_pr_number >= 0) {
          kw_pr_count_array[kw_pr_number + 1]++;
        } else {
          kw_pr_count_array.push(target_word, 1);
        }
      } else {
        if (kw_org_number >= 0) {
          kw_org_count_array[kw_org_number + 1]++;
        } else {
          kw_org_count_array.push(target_word, 1);
        }
      }

      // 全体数のカウント
      if (kw_all_number >= 0) {
        kw_all_count_array[kw_all_number + 1]++;
      } else {
        kw_all_count_array.push(target_word, 1);
      }
    }
  }

  // どの種類かの配列がnullの場合、エラーになるので、事前にnullの場合は値を入れておく
  // 二次元配列は中に[]があるのでnullではないため、lengthが０かで判断
  isNull_and_getMessage(kw_afl_count_array);
  isNull_and_getMessage(kw_pr_count_array);
  isNull_and_getMessage(kw_org_count_array);
  isNull_and_getMessage(kw_all_count_array);

  // １次元配列を２次元配列に変えている
  let kw_pr_count_array_2 = array1_to_array2_ver2(kw_pr_count_array);
  let kw_afl_count_array_2 = array1_to_array2_ver2(kw_afl_count_array);
  let kw_org_count_array_2 = array1_to_array2_ver2(kw_org_count_array);
  let kw_all_count_array_2 = array1_to_array2_ver2(kw_all_count_array);

  //outputシートの前の内容削除＆貼り付ける
  output_sheet.getRange(3, 1, output_sheet.getLastRow(), 8).clearContent();
  set_count_list_ver3(kw_pr_count_array_2, "kw_pr");
  set_count_list_ver3(kw_afl_count_array_2, "kw_afl");
  set_count_list_ver3(kw_org_count_array_2, "kw_org");
  set_count_list_ver3(kw_all_count_array_2, "kw_all");

  console.log(kw_pr_count_array_2);
  console.log(kw_org_count_array_2);
}


//---------------------------------------------------------------------------------------------------------------------


