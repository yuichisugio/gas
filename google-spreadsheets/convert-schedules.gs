// スプレッドシートのスケジュールを、一覧の指定行に入力するGAS
function autoInput_ver2() {
	// キャンペーンシートのヘッダー情報を取得
	const headers = getHeaders(campaignSheet, 2);

	// スケジュールと出力行情報の取得
	const scheduleArray = scheduleParseSheet.getDataRange().getValues();
	const outputRow = scheduleParseSheet.getRange('D1').getValue();

	// D1セルの赤枠設定
	resetCellBorder(scheduleParseSheet, 'D1');

	// 今日の日付
	const today = new Date();

	// フラグの設定
	const flags = initializeFlags(campaignSheet, headers, outputRow);

	// メニュー列の情報を取得
	const menu = getCellValueByHeader(campaignSheet, headers, outputRow, 'メニュー');

	// タスク名の対応表をオブジェクトで作成。スプシ一つの名前で一覧の二つのタスク記載がある場合は、配列にまとめている
	const taskNameMap = {
		'画像ご提出': '回収日',
		'お渡し': ['リスト事前提出', '画像提出'],
	};

	// スケジュールのタスクと日付情報を処理
	// 月のfor文
	for (let i = 0; i < 4; i++) {
		// 日のfor文
		for (let j = 0; j < scheduleArray.length; j++) {
			// このパターンの日付を作成
			const [day, month] = getDayAndMonth(scheduleArray, j, i);
			const taskDate = createTaskDate(today, month, day);

			// taskDateが日付ではない場合はスキップ
			if (isNaN(taskDate.getTime())) {
				continue;
			}

			// このパターンの日のCCとCLタスクを配列に入れる
			const taskList = [scheduleArray[j][i * 4 + 2], scheduleArray[j][i * 4 + 3]];

			// CCとCLタスクの二回分だけ繰り返して、タスクを処理
			taskList.forEach((task) => {
				// 今回の日付にタスクがある場合
				if (task) {
					processTask(task, taskDate, campaignSheet, headers, outputRow, flags, menu, taskNameMap);
				}
			});
		}
	}

	// フラグに基づく後処理
	postProcessFlags(flags, campaignSheet, headers, outputRow);
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// セルの赤枠設定をリセット
function resetCellBorder(sheet, cell) {
	sheet
		.getRange(cell)
		.setBorder(true, true, true, true, null, null, '#FF0000', SpreadsheetApp.BorderStyle.SOLID)
		.setHorizontalAlignment('center');
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// フラグの初期化
function initializeFlags(sheet, headers, row) {
	return {
		adInfo: new TeamXxxxFlagRange(sheet, headers, row, '広告配信有無'),
		draftCheckInfo: new TeamXxxxFlagRange(sheet, headers, row, 'チェック種別'),
		kolSubmitBeforeJoinInfo: new TeamXxxxFlagRange(sheet, headers, row, 'リスト事前提出'),
		clChooseKolInfo: new TeamXxxxFlagRange(sheet, headers, row, 'CL選定・開示'),
		shippingWay: new TeamXxxxFlagRange(sheet, headers, row, '発送元'),
		adSettingInfo: new TeamXxxxFlagRange(sheet, headers, row, '広告配信設定〆'),
		tieupInfo: new TeamXxxxFlagRange(sheet, headers, row, 'メニュー'),
		eventPresentationInfo: new TeamXxxxFlagRange(sheet, headers, row, 'イベント'),
		pixelTagInfo: new TeamXxxxFlagRange(sheet, headers, row, 'ピクセルタグ発行')
	};
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 列番号をヘッダー名で取得する
function getCellValueByHeader(sheet, headers, row, headerName) {
	const column = headers[headerName];
	if (!column) {
		throw new Error(`ヘッダー名 "${headerName}" が見つかりません。`);
	}
	return sheet.getRange(row, column).getValue();
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 日と月を取得する関数。createTaskDate()の比較用に数値型に変換するために必要
function getDayAndMonth(scheduleArray, rowIndex, monthIndex) {
	const day = scheduleArray[rowIndex][monthIndex * 4];
	let month = parseInt(scheduleArray[2][monthIndex * 4].replace('月', ''));

	return [day, month];
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 日付を作成する関数。「今日が6月以降」かつ「monthが1以上3以下の場合」は来年扱いにするために必要
function createTaskDate(today, month, day) {
	let referenceDay = new Date(today);

	// 「今日が6月以降」かつ「monthが1以上3以下の場合」は来年扱い
	if (referenceDay.getMonth() > 5 && month >= 1 && month <= 5) {
		referenceDay.setFullYear(referenceDay.getFullYear() + 1);
	}

	return new Date(referenceDay.getFullYear(), month - 1, day);
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// タスクを処理する関数。フラグの更新＆日付の設定のための関数
function processTask(task, taskDate, sheet, headers, row, flags, menu, taskNameMap) {
	// 広告配信開始の処理
	if (task === '広告配信開始') {
		// 広告配信の有無フラグを1更新
		flags.adInfo.flag++;
		// 「広告配信設定〆」列に「広告配信開始」の前日の日付を記載
		const adSettingDate = new Date(taskDate);
		adSettingDate.setDate(adSettingDate.getDate() - 1);
		setCellValue(sheet, row, flags.adSettingInfo.columnNumber, adSettingDate);
	}
	if (task === 'X広告審査') {
		// 広告配信の有無フラグを2更新
		flags.adInfo.flag += 2;
	}

	// 薬事チェックの種類フラグ。CCのみ→1、CLのみ→5、両者→6、事後→100
	if (task === '法務チェック') {
		flags.draftCheckInfo.flag += 1;
	}
	if (task === '原稿ご確認' || task === '初稿ご確認') {
		flags.draftCheckInfo.flag += 5;
	}
	if (task === '①投稿確認') {
		flags.draftCheckInfo.flag += 100;
	}

	// イベント・商品発表会の処理
	if (task.includes('商品発表会') || task.includes('イベント')) {
		flags.eventPresentationInfo.flag++;
		setCellValue(sheet, row, flags.eventPresentationInfo.columnNumber, taskDate);
	}

	// インフルエンサー投稿〆切の処理
	if (
		task === 'インフルエンサー投稿〆切' ||
		task === '①インフルエンサー投稿〆切②タイアップタグ設定〆切' ||
		task === '①インフルエンサー投稿〆切②タイアップタグ設定〆切・スクショ回収〆切'
	) {
		processPostingDeadline(taskDate, sheet, headers, row, flags);
	}

	// 「CL選定・開示」の処理
	if (task === 'CC一時絞り') {
		flags.clChooseKolInfo.range.setValue('CC一時絞り');
		flags.clChooseKolInfo.flag += 3;
	}
	if (task === 'オファーリスト・募集ページお渡し' || task === 'オファーリスト・オリエンページお渡し') {
		flags.clChooseKolInfo.range.setValue('事前リスト選定');
		const returnJoinedListColumn = headers['応募者リスト戻し'];
		setCellValue(sheet, row, returnJoinedListColumn, 'ー');
		flags.clChooseKolInfo.flag += 2;
		flags.kolSubmitBeforeJoinInfo.flag++;
		flags.tieupInfo.flag++;
	}
	if (task === '応募者リスト提出' || task === '応募者リストお渡し') {
		flags.clChooseKolInfo.range.setValue('募集後CL選定');
		flags.clChooseKolInfo.flag += 1;
	}

	// 発送元の処理
	if (task === '商品発送日') {
		flags.shippingWay.flag += 1;
	}
	if (task === '当選者確定＆当選者データお渡し' || task === '当選者確定日&当選者データお渡し') {
		flags.shippingWay.flag += 2;
	}

	// タイアップのフラグ設定
	if (menu.includes('タイアップ') || menu.includes('投稿必須')) {
		flags.tieupInfo.flag++;
	}

	// 広告配信のピクセルタグの有無フラグ
	if (task.includes('ピクセル')) {
		flags.pixelTagInfo.flag++;
	}

	// タスク名に基づいて対応するヘッダーを取得し、日付を設定。
	// taskキーを入れて値(列番号)が取得できる場合
	if (headers[task]) {
		setCellValue(sheet, row, headers[task], taskDate);

		// 値(列番号)が取得できない = 完全一致のタスク名がない場合
	} else if (taskNameMap[task]) {
		// 三項演算子。Array.isArray(taskNameMap[task])でtrue(=taskNameMap[task]が配列である)の場合は、targetHeadersに、taskNameMap[task]を入れて、false(=配列ではない場合)の場合は[taskNameMap[task]]を入れる
		const targetHeaders = Array.isArray(taskNameMap[task]) ? taskNameMap[task] : [taskNameMap[task]];
		targetHeaders.forEach((headerName) => {
			// ヘッダー名を指定して、列番号を取得できる場合
			if (headers[headerName]) {
				setCellValue(sheet, row, headers[headerName], taskDate);
			} else {
				console.log(`ヘッダー名 "${headerName}" が見つかりません。`);
			}
		});
	}
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 投稿〆切・報告〆切・レポート〆切の処理のための関数。メニューによってスプシの「投稿〆切」「報告〆切」の扱いが違う
function processPostingDeadline(taskDate, sheet, headers, row, flags) {
	const postDateColumn = headers['投稿〆切'];
	const informDateColumn = headers['報告〆切'];
	const reportDateColumn = headers['Excel\nレポ〆切'];

	let postDate = new Date(taskDate);
	let informDate = new Date(taskDate);
	let reportDate = new Date(taskDate);

	// タイアップの場合
	if (flags.tieupInfo.flag > 0) {
		// 投稿〆切を設定
		setCellValue(sheet, row, postDateColumn, postDate);

		// 報告〆切を設定（投稿〆切の7日後）
		informDate.setDate(informDate.getDate() + 7);
		setCellValue(sheet, row, informDateColumn, informDate);

		// レポート〆切を設定（報告〆切の3営業日後）
		reportDate = getNextBusinessDay(informDate, 3);
		setCellValue(sheet, row, reportDateColumn, reportDate);
	} else {
		// 報告〆切を設定
		setCellValue(sheet, row, informDateColumn, informDate);

		// 投稿〆切を設定（報告〆切の7日前）
		postDate.setDate(postDate.getDate() - 7);
		setCellValue(sheet, row, postDateColumn, postDate);

		// レポート〆切を設定（報告〆切の3営業日後）
		reportDate = getNextBusinessDay(informDate, 3);
		setCellValue(sheet, row, reportDateColumn, reportDate);
	}
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// フラグに基づく後処理を行う関数。フラグを見て日付以外を記載していく
function postProcessFlags(flags, sheet, headers, row) {
	// 広告配信有無の設定
	if (flags.adInfo.flag > 0) {
		// 三項演算子。２以上の場合はX配信。1ならIG第三者配信
		flags.adInfo.range.setValue(flags.adInfo.flag > 1 ? '〇(X)' : '〇(第三者)');
		if (flags.adInfo.flag === 1) {
			setCellValue(sheet, row, headers['LS広告配信対象絞込'], 'ー');
			setCellValue(sheet, row, headers['LSクリエイティブ作成'], 'ー');
			setCellValue(sheet, row, headers['KOL二次利用許諾'], 'ー');
		}
	} else {
		flags.adInfo.range.setValue('なし');
		// 広告配信がない場合、関連するセルを「ー」で埋める
		fillCells(sheet, row, flags.adInfo.columnNumber + 1, 22, 'ー');
	}

	// 薬事チェックの設定
	if (flags.draftCheckInfo.flag > 0) {
		if (flags.draftCheckInfo.flag >= 100) {
			flags.draftCheckInfo.range.setValue('事後CC');
		} else if (flags.draftCheckInfo.flag >= 6) {
			flags.draftCheckInfo.range.setValue('両者');
			fillCells(sheet, row, headers['①投稿チェック'], 6, 'ー');
		} else if (flags.draftCheckInfo.flag >= 5) {
			flags.draftCheckInfo.range.setValue('CL');
			fillCells(sheet, row, headers['①投稿チェック'], 6, 'ー');
		} else if (flags.draftCheckInfo.flag >= 1) {
			flags.draftCheckInfo.range.setValue('CC');
			fillCells(sheet, row, headers['①投稿チェック'], 6, 'ー');
		}
	} else {
		// 薬事なしの場合
		fillCells(sheet, row, flags.draftCheckInfo.columnNumber, 15, 'ー');
	}

	// 「KOLリスト事前提出」の設定
	if (flags.kolSubmitBeforeJoinInfo.flag === 0) {
		flags.kolSubmitBeforeJoinInfo.range.setValue('ー');
	}

	// CL選定・開示の設定
	if (flags.clChooseKolInfo.flag === 0) {
		flags.clChooseKolInfo.range.setValue('ー');
		setCellValue(sheet, row, headers['応募者リスト戻し'], 'ー');
	}

	// 発送元の設定。flagが0→発送なし、1→レモン、2→CL発送
	if (flags.shippingWay.flag === 0) {
		fillCells(sheet, row, flags.shippingWay.columnNumber, 6, 'ー');
	} else if (flags.shippingWay.flag === 1) {
		flags.shippingWay.range.setValue('レモン');
	} else if (flags.shippingWay.flag === 2) {
		flags.shippingWay.range.setValue('CL');
	}

	// イベント・商品発表会の設定
	if (flags.eventPresentationInfo.flag === 0) {
		flags.eventPresentationInfo.range.setValue('ー');
	}

	// ピクセルタグの有無
	if (flags.pixelTagInfo.flag === 0) {
		setCellValue(sheet, row, headers['ピクセルタグ発行'], 'ー');
		setCellValue(sheet, row, headers['ピクセルタグ設置完了'], 'ー');
	}
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// セルに値を設定する
function setCellValue(sheet, row, column, value) {
	sheet.getRange(row, column).setValue(value).setNumberFormat('M/dd');
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// 複数のセルを特定の値で埋める
function fillCells(sheet, row, startColumn, count, value) {
	const values = Array(count).fill(value);
	sheet.getRange(row, startColumn, 1, count).setValues([values]);
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー
