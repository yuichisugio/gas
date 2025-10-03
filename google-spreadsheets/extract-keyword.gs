
//---------------------------------------------------------------------------------------------------------------------


function hashtag_search_count_ver6() {

  // 追加のPR_KWを聞く関数を呼び出す
  let dialog_pr_kw = get_aditional_pr_kw();

  // キャプションの列の最終行を取得している。getlastrow()だと、投稿SNS検知の「投稿なし」に反応して無駄に取得してしまう
  let caption_lastRow = get_designation_last_row(2)

  // 2行目、１列目、データがある行まで、３列目まで取得
  let hashtag_captions_array = input_sheet.getRange(2, 1, caption_lastRow, 3).getValues();

  // for文の合計にしておきたいので、for文のスコープ外で宣言する
  // 合計数カウント
  let all_regex_count_array = [];
  // pr数カウント
  let pr_regex_count_array = [];
  // アフィリエイト数カウント
  let afl_regex_count_array = [];
  // オーガニック数カウント
  let org_regex_count_array = [];

  // 正規表現パターンを定義
  // PR判定KW
  let pr_regexPattern_reg = get_regexPattern("pr", dialog_pr_kw);
  console.log(pr_regexPattern_reg);

  // アフィ判定KW
  let afl_regexPattern_reg = get_regexPattern("afl", dialog_pr_kw);
  console.log(afl_regexPattern_reg);

  // 全部が入ったKW
  let all_regexPattern_reg = get_regexPattern("all", dialog_pr_kw);
  console.log(all_regexPattern_reg);

  // 各種の投稿件数をカウント
  let type_count_pr = ["PR投稿数", 0];
  let type_count_afl = ["アフィ投稿数", 0];
  let type_count_org = ["オーガニック投稿数", 0];
  let type_count_all = ["全種類の合計投稿数", 0];

  // 1投稿づつ取得する
  for (let s = 0; s < hashtag_captions_array.length; s++) {

    // 投稿URL
    let post_url = hashtag_captions_array[s][0];

    // キャプション
    let hashtag_caption = hashtag_captions_array[s][1];

    // 投稿SNS
    let sns = hashtag_captions_array[s][2];

    // 正規表現に当てはまるか確認。当てはまる場合は配列を返す
    // 全部が入る
    let all_regex_array = get_regexPattern_array(all_regexPattern_reg, hashtag_caption);
    // PR判定用
    let pr_regex_array = get_regexPattern_array(pr_regexPattern_reg, hashtag_caption);
    // アフィ判定
    let afl_regex_array = get_regexPattern_array(afl_regexPattern_reg, hashtag_caption);

    //判定する投稿が、aflかprかのフラグ. アフィ投稿か、PR投稿か判定している
    let afl_boolean = judge_pr_afl("afl", afl_regex_array, sns);
    let pr_boolean = judge_pr_afl("pr", pr_regex_array, sns);

    // 各種の投稿件数をカウントしている
    if (afl_boolean) {
      type_count_afl[1]++;

    } else if (pr_boolean) {
      type_count_pr[1]++;

    } else {
      // prでも、アフィでもない場合は、オーガニックをプラス
      type_count_org[1]++;
    }
    // 毎回、全ての投稿合計数を１プラス
    type_count_all[1]++;

    // 確認済みの投稿件数を出している
    console.log("確認済みの投稿件数：" + s);

    // 空欄の場合、スキップ
    if (all_regex_array == null) {
      continue;
    }

    //配列に、件数を追加
    //取り出している時に、既出（配列に入っている）ならスキップして、配列に入っていない場合は配列の中に同じ単語数だけカウントするようにする。
    for (let a = 0; a < all_regex_array.length; a++) {

      let target_word = all_regex_array[a];

      // 空欄の場合、スキップ
      if (target_word == null) {
        continue;
      }

      // それぞれが、regex_arrayにあるかどうか見て、ある場合は、そのインデックス番号、ない場合は-1を返す
      let afl_number = afl_regex_count_array.indexOf(target_word);
      let pr_number = pr_regex_count_array.indexOf(target_word);
      let organic_number = org_regex_count_array.indexOf(target_word);
      let all_number = all_regex_count_array.indexOf(target_word);

      // アフィリエイトカウント
      // 投稿SNSがXだけにしている。アフィを先にすることで、#adは、Twitterだけ先に見てアフィ判定して、残りのadはアフィではなくprだと言えるようにしている
      if (afl_boolean) {
        // その単語が何番目に入っているかのインデックス番号を返すから、配列の１番初め（0）に入っている場合は、いつになっても０なので、number>0だと、いつになっても新しく追加になってしまう。
        if (afl_number >= 0) {
          afl_regex_count_array[afl_number + 1]++;
        } else {
          afl_regex_count_array.push(target_word, 1);
        }

        // PRカウント
      } else if (pr_boolean) {
        if (pr_number >= 0) {
          pr_regex_count_array[pr_number + 1]++;
        } else {
          pr_regex_count_array.push(target_word, 1);
        }

        // オーガニックカウント
      } else {
        if (organic_number >= 0) {
          org_regex_count_array[organic_number + 1]++;
        } else {
          org_regex_count_array.push(target_word, 1);
        }
      }

      // allカウント
      if (all_number >= 0) {
        all_regex_count_array[all_number + 1]++;
      } else {
        all_regex_count_array.push(target_word, 1);
      }
    }
  }

  // 配列内に、値がない場合、エラーになるので、メッセージを代わりに入れておく。一応エラーにならないように下の表示部分でも対応している
  isNull_and_getMessage(pr_regex_count_array);
  isNull_and_getMessage(afl_regex_count_array);
  isNull_and_getMessage(org_regex_count_array);
  isNull_and_getMessage(all_regex_count_array);

  // １次元配列を２次元配列に変えている
  let pr_regex_count_array_2 = array1_to_array2_ver2(pr_regex_count_array);
  let afl_regex_count_array_2 = array1_to_array2_ver2(afl_regex_count_array);
  let organic_regex_count_array_2 = array1_to_array2_ver2(org_regex_count_array);
  let all_regex_count_array_2 = array1_to_array2_ver2(all_regex_count_array);

  // 各種投稿件数の２次元配列を作成
  let type_count_array_2 = [type_count_afl, type_count_pr, type_count_org, type_count_all];

  // outputシートの前の内容削除＆貼り付ける
  // 何も単語が当てはまらない場合、行数が0になるから、最低1行は指定してね！ってエラー出る。なので、配列が空なら表示させず、1以上ある場合は表示。Ï
  output_sheet.getRange(3, 9, output_sheet.getLastRow(), 10).clearContent();
  set_count_list_ver3(pr_regex_count_array_2, "pr");
  set_count_list_ver3(afl_regex_count_array_2, "afl");
  set_count_list_ver3(organic_regex_count_array_2, "organic");
  set_count_list_ver3(all_regex_count_array_2, "all");
  set_count_list_ver3(type_count_array_2, "type");

  // 投稿数の円グラフを作成
  make_type_graph_ver1();
}


//---------------------------------------------------------------------------------------------------------------------


