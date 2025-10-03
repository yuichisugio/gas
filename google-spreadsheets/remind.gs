// リマインド関数。これをトリガーに設定。
function getTeamXxxxOpePmKolRemind_ver5() {
	// 今日の日付を取得
	let today = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd').split('-');
	today = new Date(today[0], today[1] - 1, today[2]);

	// TEAM_XXXX一覧のデータのある一番下の行を取得
	const teamXxxxLastRow = campaignSheet.getLastRow();
	// リマインドが必要な案件の開始行を指定
	const startRow = 2418;
	// 必要な範囲を計算
	const searchRow = teamXxxxLastRow - startRow;

	// TEAM_XXXX一覧のヘッダー名とカラム番号が入った配列を取得
	// 最終列を取得
	const teamXxxxLastColumn = campaignSheet.getLastColumn();
	// ヘッダー行をすべて（最終列まで）取得
	let headers = campaignSheet.getRange(2, 1, 1, teamXxxxLastColumn).getValues()[0];

	// リマインドする際に送る情報を取得。1列ずつ取得するので２次元配列を１次元配列に修正。
	const menu = campaignSheet
		.getRange(startRow, headers.indexOf('メニュー') + 1, searchRow)
		.getValues()
		.flat(); // メニュー
	const campaignId = campaignSheet
		.getRange(startRow, headers.indexOf('ID') + 1, searchRow)
		.getValues()
		.flat(); // キャンペーンID
	const companyName = campaignSheet
		.getRange(startRow, headers.indexOf('会社名') + 1, searchRow)
		.getValues()
		.flat(); // 会社名
	const products = campaignSheet
		.getRange(startRow, headers.indexOf('商品名') + 1, searchRow)
		.getValues()
		.flat(); // 商品名
	const pmMember = campaignSheet
		.getRange(startRow, headers.indexOf('PM') + 1, searchRow)
		.getValues()
		.flat(); //pm担当者
	const opeMember = campaignSheet
		.getRange(startRow, headers.indexOf('OPE') + 1, searchRow)
		.getValues()
		.flat(); //ope担当者
	const kolMember = campaignSheet
		.getRange(startRow, headers.indexOf('KOL') + 1, searchRow)
		.getValues()
		.flat(); //kol担当者
	const status = campaignSheet
		.getRange(startRow, headers.indexOf('状況') + 1, searchRow)
		.getValues()
		.flat(); //ステータス
	const campaignPage = campaignSheet
		.getRange(startRow, headers.indexOf('概要') + 1, searchRow)
		.getValues()
		.flat(); //CP概要

	const lpDeadLine = campaignSheet
		.getRange(startRow, headers.indexOf('LP提出') + 1, searchRow)
		.getValues()
		.flat(); // LP提出
	const joinStart = campaignSheet
		.getRange(startRow, headers.indexOf('募集開始') + 1, searchRow)
		.getValues()
		.flat(); //募集開始
	const deliveryDeadLine = campaignSheet
		.getRange(startRow, headers.indexOf('納品期限') + 1, searchRow)
		.getValues()
		.flat(); // 納品期限日
	const joinDeadLine = campaignSheet
		.getRange(startRow, headers.indexOf('募集〆切') + 1, searchRow)
		.getValues()
		.flat(); // 募集〆切
	const joinedRule = campaignSheet
		.getRange(startRow, headers.indexOf('CL選定・開示') + 1, searchRow)
		.getValues()
		.flat(); // CL選定・開示
	const kolFix = campaignSheet
		.getRange(startRow, headers.indexOf('当選者確定') + 1, searchRow)
		.getValues()
		.flat(); //当選者確定
	const startDraftSubmit = campaignSheet
		.getRange(startRow, headers.indexOf('下書き開始') + 1, searchRow)
		.getValues()
		.flat(); //下書き開始
	const draftDeadLine = campaignSheet
		.getRange(startRow, headers.indexOf('下書き〆') + 1, searchRow)
		.getValues()
		.flat(); //下書き〆
	const postDeadLine = campaignSheet
		.getRange(startRow, headers.indexOf('投稿〆切') + 1, searchRow)
		.getValues()
		.flat(); // 投稿〆切

	/*
  各担当者ごとの送る文章を設定。for文で初期化したく無いので、外に設定
  PM→test-1、test-2
  OPE→test-1、test-2
  KOL→test-1、test-2
  一覧のPM担当者と同じ名前にする必要があるので注意！メンバーが増えた時は、ここだけ追加すればOK！
  */
	// PM
	let pmTest1Info = new TeamXxxxMember('test-1', 'UXXXXXXXXXXX');
	let pmTest2Info = new TeamXxxxMember('test-2', 'UXXXXXXXXXXX');
	let pmMemberList = [pmTest1Info, pmTest2Info];
	// OPE
	let opeTest1Info = new TeamXxxxMember('test-1', 'UXXXXXXXXXXX');
	let opeTest2Info = new TeamXxxxMember('test-2', 'UXXXXXXXXXXX');
	let opeMemberList = [opeTest1Info, opeTest2Info];
	// KOL
	let kolTest1Info = new TeamXxxxMember('test-1', 'U07LXXXXXX');
	let kolTest2Info = new TeamXxxxMember('test-2', 'U07LXXXXXX');
	let kolMemberList = [kolTest1Info, kolTest2Info];

	let i = 0;
	// 1行づつ探していく
	for (let id of campaignId) {
		// ステータスが「未受注」または「report済」の行はスキップ
		if (status[i] == '未受注' || status[i] == 'report済' || status[i] == '失注' || status[i] == '完了') {
			i++;
			continue;
		}

		// リマインドタスクのテキストを追加する関数。スコープ内変数を使用したいため、スコープ内に関数を定義
		function addText(assignedMember, taskText, team) {
			if (team === 'pm') {
				for (let i = 0; i < pmMemberList.length; i++) {
					if (assignedMember.includes(pmMemberList[i].memberName)) {
						pmMemberList[i].sendText += templateText + taskText + '\n\n';
						pmMemberList[i].flag++;
					}
				}
			} else if (team === 'ope') {
				for (let i = 0; i < opeMemberList.length; i++) {
					if (assignedMember.includes(opeMemberList[i].memberName)) {
						opeMemberList[i].sendText += templateText + taskText + '\n\n';
						opeMemberList[i].flag++;
					}
				}
			} else if (team === 'kol') {
				for (let i = 0; i < kolMemberList.length; i++) {
					if (assignedMember.includes(kolMemberList[i].memberName)) {
						kolMemberList[i].sendText += templateText + taskText + '\n\n';
						kolMemberList[i].flag++;
					}
				}
			}
		}

		// 送るテキストの初期値を設定
		let templateText =
			`\nキャンペーンID：<${campaignPage[i]}|` +
			id +
			'>\nメニュー：' +
			menu[i] +
			'\n会社名：' +
			companyName[i].replace(/[\r\n]+/g, ' ') +
			'\n商品名：' +
			products[i].replace(/[\r\n]+/g, ' ') +
			'\n';

		/*
    PMリマインドタスク→「LP提出」「納品期日」「応募者リスト提出」「初稿ご確認（LS法務チェック〆）」「広告配信投稿一覧提出」「ブースト依頼」
    OPEリマインド→「LP提出」「予約公開日」「募集開始」「納品期限日」「当選者確定期日」「レポ〆切」
    */

		// LP提出。PMとOPE両方
		if (lpDeadLine[i].toString() == today.toString()) {
			let lpDeadLineText = '*今日、提出日だよ！*';
			addText(pmMember[i], lpDeadLineText, 'pm');
			addText(opeMember[i], lpDeadLineText, 'ope');
		}

		// 予約公開日。OPEのみ
		// 予約公開日の日付を作成。募集開始の1日前
		let reservation = new Date(joinStart[i]);
		reservation.setDate(reservation.getDate() - 1);
		if (reservation.toString() == today.toString()) {
			let reservationText = '*今日、予約公開日だよ！*\n※募集開始日の1日前';
			addText(opeMember[i], reservationText, 'ope');
		}

		// 募集開始。OPEのみ
		if (joinStart[i].toString() == today.toString()) {
			let joinStartText = '*今日、募集開始日だよ！*';
			addText(opeMember[i], joinStartText, 'ope');
		}

		// 納品期限。PMとOPE両方
		if (deliveryDeadLine[i].toString() == today.toString()) {
			let deliveryDeadLineText = '*今日、納品期日だよ！*';
			addText(pmMember[i], deliveryDeadLineText, 'pm');
			addText(opeMember[i], deliveryDeadLineText, 'ope');
		}

		// 募集〆切の3営業日前。KOLのみ
		if (getPreviousBusinessDay(joinDeadLine[i], 3).toString() == today.toString()) {
			let joinDeadLineText = '*今日、募集〆切の3営業日前だよ！*';
			addText(kolMember[i], joinDeadLineText, 'kol');
		}

		// 応募者リスト提出。PMのみ。募集〆切の1営業日後にリマインド。条件は「CL選定・開示」の列が「募集後CL選定」or「CC一時絞り」の場合のみリマインド。
		if (
			getNextBusinessDay(joinDeadLine[i], 1).toString() == today.toString() &&
			(joinedRule[i] == '募集後CL選定' || joinedRule[i] == 'CC一時絞り')
		) {
			let submitListText = '*今日、応募者リスト提出日だよ！*';
			addText(pmMember[i], submitListText, 'pm');
		}

		// 当選者確定日の1営業日前。KOLのみ
		if (getPreviousBusinessDay(kolFix[i], 1).toString() == today.toString()) {
			let kolFixBeforeText = '*今日、当選者確定日の1営業日前だよ！*';
			addText(kolMember[i], kolFixBeforeText, 'kol');
		}

		// 当選者確定。OPEのみ
		if (kolFix[i].toString() == today.toString()) {
			let kolFixText = '*今日、当選者確定日だよ！*';
			addText(opeMember[i], kolFixText, 'ope');
		}

		// 下書き開始当日（薬事ありギフ）、土日開始の場合は金曜。KOLのみ
		if (getBusinessDay(startDraftSubmit[i]).toString() == today.toString() && menu[i] == 'ギフティング') {
			let startDraftSubmitText = '*今日、下書き開始当日だよ！';
			addText(kolMember[i], startDraftSubmitText, 'kol');
		}

		// iに1足して、TEAM_XXXX一覧の次の行に進む
		i++;
	}

	// リマインドする案件がなければ違う文言にする
	let noRemindText = '\n*今日のリマインドタスクなし*';

	// リマインドをしない条件は、「土日」かつ「リマインドする案件がない」場合。メンバー分だけ繰り返し処理
	// PM
	for (let co = 0; co < pmMemberList.length; co++) {
		flagToText(pmMemberList[co].flag, pmMemberList[co].sendText, 'pm', noRemindText, today);
	}
	// OPE
	for (let co = 0; co < opeMemberList.length; co++) {
		flagToText(opeMemberList[co].flag, opeMemberList[co].sendText, 'ope', noRemindText, today);
	}
	// KOL
	for (let co = 0; co < kolMemberList.length; co++) {
		flagToText(kolMemberList[co].flag, kolMemberList[co].sendText, 'kol', noRemindText, today);
	}
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// TeamXxxxチームの各メンバーの、各フラグ・テキスト・名前を保存する。
class TeamXxxxMember {
	constructor(memberName, mention) {
		// 送る案件があるかのフラグ
		this.flag = 0;
		// 担当者名
		this.memberName = memberName;
		// 送る文章
		this.sendText = `<@${mention}>\n*${memberName}さん担当のリマインド！*\n`;
	}
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 休日(土日祝日)で案件がない場合はリマインド投稿を飛ばさない関数
function flagToText(flag, text, team, noRemindText, today) {
	// リマインドする案件が場合。
	if (flag != 0) {
		sendSlack(text, team);

		// 「土日・祝日以外」かつ「リマインドする案件がない」場合。
	} else if (flag == 0 && today.getDay() != 0 && today.getDay() != 6 && !isHoliday(today)) {
		text += noRemindText;
		sendSlack(text, team);
	}
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 引数が休日なら1営業日前の平日を返し、平日なら引数の日を返す関数
function getBusinessDay(startDate) {
	// 引数の日付を複製
	let targetDate = startDate ? new Date(startDate) : new Date();

	// 休日の場合は1営業日前にする
	// 土曜or日曜or祝日の場合
	if (targetDate.getDay() === 0 || targetDate.getDay() === 6 || isHoliday(targetDate)) {
		do {
			// targetDateを1日前にする
			targetDate.setDate(targetDate.getDate() - 1);
			// 1日前にした日が土曜or日曜or祝日の場合は何度も1日前にする
		} while (targetDate.getDay() === 0 || targetDate.getDay() === 6 || isHoliday(targetDate));
	}

	return targetDate;
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー
