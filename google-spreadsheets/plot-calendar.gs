// 祝日か判定する関数
function isJpHoliday(date) {
	// 祝日の配列。毎年手入力が必要
	let holidays = [
		'2024-01-01',
		'2024-01-02',
		'2024-01-03',
		'2024-01-08',
		'2024-02-11',
		'2024-02-12',
		'2024-02-23',
		'2024-03-20',
		'2024-04-29',
		'2024-05-03',
		'2024-05-04',
		'2024-05-05',
		'2024-05-06',
		'2024-07-15',
		'2024-08-11',
		'2024-08-12',
		'2024-09-16',
		'2024-09-22',
		'2024-09-23',
		'2024-10-14',
		'2024-11-03',
		'2024-11-04',
		'2024-11-23',
		'2024-12-31',
		'2025-01-01',
		'2025-01-02',
		'2025-01-03',
		'2025-01-13',
		'2025-02-11',
		'2025-02-23',
		'2025-02-24',
		'2025-03-20',
		'2025-04-29',
		'2025-05-03',
		'2025-05-04',
		'2025-05-05',
		'2025-05-06',
		'2025-07-21',
		'2025-08-11',
		'2025-09-15',
		'2025-09-23',
		'2025-10-13',
		'2025-11-03',
		'2025-11-23',
		'2025-11-24',
		'2025-12-31',
		'2026-01-01',
		'2026-01-02',
		'2026-01-03',
		'2026-01-12',
		'2026-02-11',
		'2026-02-23',
		'2026-03-20',
		'2026-04-29',
		'2026-05-03',
		'2026-05-04',
		'2026-05-05',
		'2026-05-06',
		'2026-07-20',
		'2026-08-11',
		'2026-09-21',
		'2026-09-22',
		'2026-09-23',
		'2026-10-12',
		'2026-11-03',
		'2026-11-23',
		'2026-12-31',
		'2027-01-01',
		'2027-01-02',
		'2027-01-03',
		'2027-01-11',
		'2027-02-11',
		'2027-02-23',
		'2027-03-21',
		'2027-03-22',
		'2027-04-29',
		'2027-05-03',
		'2027-05-04',
		'2027-05-05',
		'2027-07-19',
		'2027-08-11',
		'2027-09-20',
		'2027-09-23',
		'2027-10-11',
		'2027-11-03',
		'2027-11-23',
		'2027-12-31',
		'2028-01-01',
		'2028-01-02',
		'2028-01-03',
		'2028-01-10',
		'2028-02-11',
		'2028-02-23',
		'2028-03-20',
		'2028-04-29',
		'2028-05-03',
		'2028-05-04',
		'2028-05-05',
		'2028-07-17',
		'2028-08-11',
		'2028-09-18',
		'2028-09-22',
		'2028-10-09',
		'2028-11-03',
		'2028-11-23',
		'2028-12-31',
		'2029-01-01',
		'2029-01-02',
		'2029-01-03',
		'2029-01-08',
		'2029-02-11',
		'2029-02-12',
		'2029-02-23',
		'2029-03-20',
		'2029-04-29',
		'2029-04-30',
		'2029-05-03',
		'2029-05-04',
		'2029-05-05',
		'2029-07-16',
		'2029-08-11',
		'2029-09-17',
		'2029-09-23',
		'2029-09-24',
		'2029-10-08',
		'2029-11-03',
		'2029-11-23',
		'2029-12-31',
		'2024-12-29',
		'2024-12-30',
		'2025-12-29',
		'2025-12-30',
		'2026-12-29',
		'2026-12-30'
	];

	// 引数の日付の形式を整える
	let formattedDate = Utilities.formatDate(date, 'JST', 'yyyy-MM-dd');

	/* 
  toISOString()をすると、JSTではなくUTCで判断されて、formattedDateに日本時間の9時間前の時間(前日)が入ってしまうので、背景色青の部分が祝日の場合1日ずれている。
  toISOString()の代わりに、Utilities.formatDate(date,"JST","yyyy-MM-dd"))で記載する
  ISOでyyyy-MM-dd形式にしているのを、utilitiesでその形式に変更すると、タイムゾーンがずれずに行けそう
  let formattedDate = date.toISOString().split('T')[0];
  */

	// holidays配列にある日付であれば、trueを返す
	return holidays.includes(formattedDate);
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 「作成」ボタンで起動する関数
function getSchedule() {
	// スプシを取得
	const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
	const inputSheet = spreadSheet.getSheetByName('INPUT');
	const outputSheet = spreadSheet.getSheetByName('OUTPUT');

	// 「OUTPUT」タブの前の内容を削除
	outputSheet.clear();

	// メニューと開始日を「INPUT」から取得
	const campaignType = inputSheet.getRange('B1').getValue();
	let startDate = inputSheet.getRange('D1').getValue();
	startDate = new Date(startDate);

	//「範囲」セルに指定した範囲のそれぞれのタスクを取得している。配列になっている。
	let events = inputSheet.getRange(campaignType).getValues();

	// 「OUTPUT」タブにスケを表示するための関数を呼び出す
	plotCalendar(startDate, outputSheet, events);
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 「OUTPUT」タブにスケを表示するための関数
function plotCalendar(startDate, outputSheet, events) {
	// OUTPUTに記載するタスクの開始日（オリシ回収）まで、スキップするための日付を作成。比較用
	const referenceDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());

	// 曜日の配列。getdayでは0-6の数字のため
	const jst = ['日', '月', '火', '水', '木', '金', '土'];

	// タスクの数カウントの初期値
	// タスクの合計値
	let eventNum = 0;
	// タスクの数のカウント
	let eventCount = 0;

	outputSheet.getRange(1, 1).setValue('CompanyXxxx TeamXxxx Squareスケジュール');

	// 月の記載。1月毎に記載するfor分、4ヶ月分表示
	for (let l = 0; l < 4; l++) {
		// OUTPUTカレンダーに記載するために、INPUT開始日の開始月の1日目を取得
		let monthStartDate = new Date(startDate.getFullYear(), startDate.getMonth() + l, 1);

		// 下で日を足していった時に、次の月になっているか判断する用。今月を入れておき、今月の間は日を足していく。
		let month = startDate.getMonth() + l;

		// for文で月を足して行った時に、13月、14月、15月になるので、1月、2月、3月に修正。月は0-11なので1つ少ない
		if (startDate.getMonth() + l == 12) {
			monthStartDate = new Date(startDate.getFullYear() + 1, 0, 1);
			month = 0;
		} else if (startDate.getMonth() + l == 13) {
			monthStartDate = new Date(startDate.getFullYear() + 1, 1, 1);
			month = 1;
		} else if (startDate.getMonth() + l == 14) {
			monthStartDate = new Date(startDate.getFullYear() + 1, 2, 1);
			month = 2;
		}

		// この月のタスクを記載する列数を指定
		let rowNum = l * 4;

		// タスクより上の部分、ヘッダーを記載
		outputSheet
			.getRange(3, 1 + rowNum, 1, 2)
			.merge()
			.setValue(monthStartDate.getMonth() + 1 + '月')
			.setHorizontalAlignment('center')
			.setBorder(true, true, true, true, true, true);
		outputSheet.getRange(3, 1 + rowNum).setBackground('#c4c7cc');
		outputSheet
			.getRange(4, 1 + rowNum)
			.setValue('日')
			.setHorizontalAlignment('center');
		outputSheet.getRange(4, 1 + rowNum, 1, 4).setBorder(true, true, true, true, true, true);
		outputSheet
			.getRange(4, 2 + rowNum)
			.setValue('曜')
			.setHorizontalAlignment('center');
		outputSheet
			.getRange(4, 3 + rowNum)
			.setValue('CompanyXxxx')
			.setHorizontalAlignment('center');
		outputSheet.getRange(4, 3 + rowNum).setBackground('#FFFF00');
		outputSheet
			.getRange(4, 4 + rowNum)
			.setValue('クライアント様')
			.setHorizontalAlignment('center');

		// 日の記載。1日づつfor文で回して、条件はmonthStartDateが次の月になるまで行う。monthStartDateは下の方で増やしている
		for (let i = 5; monthStartDate.getMonth() == month; i++) {
			// 日・曜日を取得
			let date = monthStartDate.getDate();
			let day = monthStartDate.getDay();
			let jstDay = jst[day];

			// 祝日か判定用に、タスクを記載する日(monthStartDate)を複製
			let holidayDate = new Date(monthStartDate.getFullYear(), monthStartDate.getMonth(), monthStartDate.getDate());
			// let holiday = calendar.getEventsForDay(holidayDate);

			// 日曜日・土曜日・祝日かを判定して、祝日なら青色にする & 休日フラグ（holidayFlag）を1にする
			// if (day == 0 || day == 6 || holiday.length > 0) {
			if (day == 0 || day == 6 || isJpHoliday(holidayDate)) {
				holidayFlag = 1;
				outputSheet.getRange(i, 1 + rowNum, 1, 4).setBackground('#bbcff0');
			} else {
				holidayFlag = 0;
			}

			// タスクを記載する日の日数や曜日を記載
			outputSheet
				.getRange(i, 1 + rowNum)
				.setValue(date)
				.setHorizontalAlignment('center');
			outputSheet.getRange(i, 1 + rowNum, 1, 4).setBorder(true, true, true, true, true, true);
			outputSheet
				.getRange(i, 2 + rowNum)
				.setValue(jstDay)
				.setHorizontalAlignment('center');

			// OUTPUTに記載する日を、1日足して次に進める
			monthStartDate.setDate(monthStartDate.getDate() + 1);

			// OUTPUTに記載するタスクが、記入した開始日より後になるまでスキップする
			if (monthStartDate <= referenceDate) {
				continue;
			}

			// タスクの記載。CCとCLのタスクを記載していくfor文。タスクの数だけ回す
			for (let m = eventNum; m < events.length; m++) {
				// 応募〆切を日曜日にずらす。前タスクから開ける日数がeventCountと同じ　＆＆　営業日でカウントするタスク　＆＆　タスク内容が「募集〆切」or「投稿〆切」or「下書き提出〆切」の時
				if (
					events[m][1] == eventCount &&
					events[m][2] == 1 &&
					(events[m][0] == '募集〆切' || events[m][0] == '投稿〆切' || events[m][0] == '提出〆切')
				) {
					// 金曜日の場合。2つ下の行(日曜日)に、記載
					if (day == 5) {
						outputSheet.getRange(i + 2, 3 + rowNum).setValue(events[m][0]);
						// plotSalesInfomation(salesSheet, events[m][0], holidayDate.setDate(holidayDate.getDate() + 2), passedDays+2);
						// 土曜日の場合。1つ下の行(日曜日)に、記載
					} else if (day == 6) {
						outputSheet.getRange(i + 1, 3 + rowNum).setValue(events[m][0]);
						// plotSalesInfomation(salesSheet, events[m][0], monthStartDate, passedDays+1);
						// それ以外の場合
					} else {
						outputSheet.getRange(i, 3 + rowNum).setValue(events[m][0]);
						// plotSalesInfomation(salesSheet, events[m][0], holidayDate, passedDays);
					}
					// 前タスクから開ける日数を0にリセット。リセットしないと何日空ける。という設定で日数記載しているのに違う内容になる
					eventCount = 0;
					// イベント合計数に１を足して、for文で最後が完了したらちゃんと終われるようにする
					eventNum += 1;

					// 応募〆切・投稿〆切以外で、前タスクから開ける日数がeventCountと同じ　&&　休日もOKのタスクの場合
				} else if (events[m][1] == eventCount && events[m][2] == 1) {
					// Chanタスクの場合
					if (events[m][3] == 0) {
						outputSheet.getRange(i, 3 + rowNum).setValue(events[m][0]);
						// CLタスクの場合
					} else {
						outputSheet.getRange(i, 4 + rowNum).setValue(events[m][0]);
					}
					eventCount = 0;
					eventNum += 1;
					// plotSalesInfomation(salesSheet, events[m][0], holidayDate, passedDays);

					// 前タスクから開ける日数がeventCountと同じ　&&　休日NGのタスク &&　この日が休日ではない場合
				} else if (events[m][1] == eventCount && holidayFlag == 0) {
					// 「キックオフ候補日」or「イベント日」の場合は、結合させる
					if (events[m][0] == 'キックオフ候補日' || events[m][0] == 'イベント日') {
						outputSheet
							.getRange(i, 3 + rowNum, 1, 2)
							.merge()
							.setValue(events[m][0])
							.setHorizontalAlignment('center');
						// ↑以外で、CCタスクの場合
					} else if (events[m][3] == 0) {
						outputSheet.getRange(i, 3 + rowNum).setValue(events[m][0]);
						// ↑以外で、CLタスクの場合
					} else {
						outputSheet.getRange(i, 4 + rowNum).setValue(events[m][0]);
					}
					eventCount = 0;
					eventNum += 1;
					// plotSalesInfomation(salesSheet, events[m][0], holidayDate, passedDays);

					// 前タスクから開ける日数がeventCountと同じ　&&　休日NGのタスク &&　この日が休日の場合
				} else if (events[m][1] == eventCount && holidayFlag == 1) {
					break;

					//eventCount（前のタスクからの空き日数）が、記載の数字まで達していなければ、eventCountを足してから、達するまでこのif文を中止して、一個外のif文に戻る
				} else {
					eventCount += 1;
					break;
				}
			}
		}
	}
	// 最後に全体を垂直報告で中央寄せ
	outputSheet.getRange(1, 1, 37, 20).setVerticalAlignment('middle');
}
