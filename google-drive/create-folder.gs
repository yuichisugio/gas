// 会社名を検索するバージョン！
function search_clientFolder_ver1() {
	// 検索ワードを取得
	let toSearch_companyName = makeFolderSheet.getRange(2, 9).getValue();

	// 検索ワードのパラメーターを作成 & そのワードを含む会社フォルダを取得。containsの後は、追加で引用符が必要
	let params = `title contains "${toSearch_companyName}"`;
	let searchNameCollection = clientFolder.searchFolders(params);

	// 検索に該当した内容が入る二次元配列
	let matchNameArray = [];

	// 検索結果に該当する会社フォルダの名前と会社フォルダURLを配列に入れる！
	while (searchNameCollection.hasNext()) {
		// フォルダを取得
		let folder = searchNameCollection.next();
		// フォルダURLを取得
		let folderUrl = folder.getUrl();
		// フォルダ名を取得
		let folderName = folder.getName();
		// 各情報を配列にプッシュ！
		matchNameArray.push([folderName, folderUrl]);
	}

	// アウトプット結果の範囲をクリア
	makeFolderSheet.getRange(2, 10, makeFolderSheet.getMaxRows() - 1, 2).clearContent();

	// 配列に値があるなら配列内容を、ないならその文言を表示！
	if (matchNameArray.length > 0) {
		makeFolderSheet.getRange(2, 10, matchNameArray.length, matchNameArray[0].length).setValues(matchNameArray);
	} else {
		makeFolderSheet.getRange(2, 10).setValue('該当する会社名なし！');
	}
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// クライアントフォルダを作成する！
function make_clientFolder_ver1() {
	// 各値を取得
	let toMake_companyName = makeFolderSheet.getRange(2, 3).getValue();
	let toMake_bradName = makeFolderSheet.getRange(2, 4).getValue();
	let toMake_goodsName = makeFolderSheet.getRange(2, 5).getValue();
	let toMake_snsName = makeFolderSheet.getRange(2, 6).getValue();
	let toMake_planName = makeFolderSheet.getRange(2, 7).getValue();

	// それぞれの階層のフォルダを入れる変数を宣言
	let companyFolder;
	let brandFolder;
	let goodsFolder;
	let reportFolder;
	let imageFolder;
	let workFolder;
	let lpImageFolder;
	let brandImageFolder;
	let otherImageFolder;
	let ui = SpreadsheetApp.getUi();

	// 会社名の入力がない場合は、入力依頼。会社名を入力しないと、新しい名前で作成する際、「new folder」という名前で作成されてしまう。
	if (toMake_companyName == '' || toMake_companyName == null || toMake_companyName == undefined) {
		ui.alert('↓のセルの下に、入力してね！\n会社名');
		return;
	}

	// ブランド名の入力がない場合は、入力依頼。ブランド名を入力しないと、新しい名前で作成する際、「new folder」という名前で作成されてしまう。
	if (toMake_bradName == '' || toMake_bradName == null || toMake_bradName == undefined) {
		ui.alert('↓のセルの下に、入力してね！\nブランド名');
		return;
	}

	// 会社フォルダがあるか確認。ないなら作成
	// 完全一致で会社名のフォルダがあるか確認
	let companyCollection = clientFolder.getFoldersByName(toMake_companyName);
	// hasNestで、コレクションの中身があるかどうかbooleanで判定
	let companyName_boolean = companyCollection.hasNext();
	if (companyName_boolean) {
		// ある場合は、そのフォルダを取得
		companyFolder = companyCollection.next();
	} else {
		// ない場合は新規で作成
		companyFolder = clientFolder.createFolder(toMake_companyName);
	}

	// ブランドフォルダがあるか確認。ないなら作成
	let brandCollection = companyFolder.getFoldersByName(toMake_bradName);
	let brandName_boolean = brandCollection.hasNext();
	if (brandName_boolean) {
		brandFolder = brandCollection.next();
	} else {
		brandFolder = companyFolder.createFolder(toMake_bradName);
	}

	// 商品名のフォルダと、その中にフォルダ群を作成
	let goods_title = 'CP_' + toMake_goodsName + '_' + toMake_snsName + toMake_planName;
	goodsFolder = brandFolder.createFolder(goods_title);
	reportFolder = goodsFolder.createFolder('レポート');
	imageFolder = goodsFolder.createFolder('画像');
	workFolder = reportFolder.createFolder('作業用');
	lpImageFolder = imageFolder.createFolder('募集画像');
	brandImageFolder = imageFolder.createFolder('ブランド画像');
	otherImageFolder = imageFolder.createFolder('素材');

	// 作成したドライブURLを表示させる！
	makeFolderSheet.getRange(2, 8).clearContent();
	let goodsFolderUrl = goodsFolder.getUrl();
	makeFolderSheet.getRange(2, 8).setValue(goodsFolderUrl);
}

// ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

// ブランド分析のフォルダを作成。ボタンを押した時は、直接↓が起動
function brandAnalysis_makeFolder_duplicateSheet_ver2() {
	// 入力名
	let analysisName = makeFolderSheet.getRange(2, 1).getValue();

	// 今日の日付を取得＆フォーマットを作成
	let formatDate = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd');

	// フォルダやシートにつける名前を作る
	let title = formatDate + '_' + analysisName;

	// ブランド分析フォルダの中に、指定した名前のフォルダを作成　＆　そのフォルダ内に作業用フォルダとテンプレ複製シートを作成
	let companyNameFolder = brandAnalysis_Folder.createFolder(title);
	companyNameFolder.createFolder('作業用');
	brandAnalysis_templateSheet.makeCopy(title, companyNameFolder);

	// 作成したドライブURLを表示させる！
	makeFolderSheet.getRange(2, 2).clearContent();
	let companyNameFolderUrl = companyNameFolder.getUrl();
	makeFolderSheet.getRange(2, 2).setValue(companyNameFolderUrl);
}

// ブランド分析のフォルダを作成。ボタンを押した時は、直接↓が起動
function brandAnalysis_makeFolder_duplicateSheet_ver2(name) {
	// ダイアログから入力もできるし、セルで入力も可能にする
	// 入力名
	let analysisName = makeFolderSheet.getRange(2, 1).getValue();
	name = analysisName;

	// 今日の日付を取得＆フォーマットを作成
	let today = new Date();
	let formatDate = Utilities.formatDate(today, 'JST', 'yyyyMMdd');

	// フォルダやシートにつける名前を作る
	let title = formatDate + '_' + name;

	// ブランド分析フォルダの中に、指定した名前のフォルダを作成　＆　そのフォルダ内に作業用フォルダとテンプレ複製シートを作成
	let companyNameFolder = brandAnalysis_Folder.createFolder(title);
	companyNameFolder.createFolder('作業用');
	brandAnalysis_templateSheet.makeCopy(title, companyNameFolder);

	// 作成したドライブURLを表示させる！
	makeFolderSheet.getRange(2, 2).clearContent();
	let companyNameFolderUrl = companyNameFolder.getUrl();
	makeFolderSheet.getRange(2, 2).setValue(companyNameFolderUrl);
}

//---------------------------------------------------------------------------------------------------------------------

// 会社名を検索するバージョン！
function search_clientFolder_ver1() {
	// 検索ワードを取得
	let toSearch_companyName = makeFolderSheet.getRange(2, 9).getValue();

	// 検索ワードのパラメーターを作成 & そのワードを含む会社フォルダを取得
	// containsの後は、追加で引用符が必要
	let params = `title contains "${toSearch_companyName}"`;
	let searchNameCollection = clientFolder.searchFolders(params);

	// 検索に該当した内容が入る二次元配列
	let matchNameArray = [];

	// 検索結果に該当する会社フォルダの名前と会社フォルダURLを配列に入れる！
	while (searchNameCollection.hasNext()) {
		// フォルダを取得
		let folder = searchNameCollection.next();
		// フォルダURLを取得
		let folderUrl = folder.getUrl();
		// フォルダ名を取得
		let folderName = folder.getName();
		// 各情報を配列にプッシュ！
		matchNameArray.push([folderName, folderUrl]);
	}

	// アウトプット結果の範囲をクリア
	makeFolderSheet.getRange(2, 10, makeFolderSheet.getMaxRows() - 1, 2).clearContent();

	// 配列に値があるなら配列内容を、ないならその文言を表示！
	if (matchNameArray.length > 0) {
		makeFolderSheet.getRange(2, 10, matchNameArray.length, matchNameArray[0].length).setValues(matchNameArray);
	} else {
		makeFolderSheet.getRange(2, 10).setValue('該当する会社名なし！');
	}
}

//---------------------------------------------------------------------------------------------------------------------

// クライアントフォルダを作成する！
function make_clientFolder_ver1() {
	// 各値を取得
	let toMake_companyName = makeFolderSheet.getRange(2, 3).getValue();
	let toMake_bradName = makeFolderSheet.getRange(2, 4).getValue();
	let toMake_goodsName = makeFolderSheet.getRange(2, 5).getValue();
	let toMake_snsName = makeFolderSheet.getRange(2, 6).getValue();
	let toMake_planName = makeFolderSheet.getRange(2, 7).getValue();

	// それぞれの階層のフォルダを入れる変数を宣言
	let companyFolder;
	let brandFolder;
	let goodsFolder;
	let reportFolder;
	let imageFolder;
	let workFolder;
	let lpImageFolder;
	let brandImageFolder;
	let otherImageFolder;
	let ui = SpreadsheetApp.getUi();

	// 会社名の入力がない場合は、入力依頼。会社名を入力しないと、新しい名前で作成する際、「new folder」という名前で作成されてしまう。
	if (toMake_companyName == '' || toMake_companyName == null || toMake_companyName == undefined) {
		ui.alert('↓のセルの下に、入力してね！\n会社名');
		return;
	}

	// ブランド名の入力がない場合は、入力依頼。ブランド名を入力しないと、新しい名前で作成する際、「new folder」という名前で作成されてしまう。
	if (toMake_bradName == '' || toMake_bradName == null || toMake_bradName == undefined) {
		ui.alert('↓のセルの下に、入力してね！\nブランド名');
		return;
	}

	// 商品名の入力がない場合は、入力依頼。商品名を入力しないと、新しい名前で作成する際、「new folder」という名前で作成されてしまう。
	if (toMake_goodsName == '' || toMake_goodsName == null || toMake_goodsName == undefined) {
		toMake_goodsName = '未定';
	}

	// 会社フォルダがあるか確認。ないなら作成
	let companyCollection = clientFolder.getFoldersByName(toMake_companyName);
	let companyName_boolean = companyCollection.hasNext();
	if (companyName_boolean) {
		companyFolder = companyCollection.next();
	} else {
		companyFolder = clientFolder.createFolder(toMake_companyName);
	}

	// ブランドフォルダがあるか確認。ないなら作成
	let brandCollection = companyFolder.getFoldersByName(toMake_bradName);
	let brandName_boolean = brandCollection.hasNext();
	if (brandName_boolean) {
		brandFolder = brandCollection.next();
	} else {
		brandFolder = companyFolder.createFolder(toMake_bradName);
	}

	// 商品名のフォルダと、その中にフォルダ群を作成
	let goods_title = 'CP_' + toMake_goodsName + '_' + toMake_snsName + toMake_planName;
	goodsFolder = brandFolder.createFolder(goods_title);
	reportFolder = goodsFolder.createFolder('レポート');
	imageFolder = goodsFolder.createFolder('画像');
	workFolder = reportFolder.createFolder('作業用');
	lpImageFolder = imageFolder.createFolder('募集画像');
	brandImageFolder = imageFolder.createFolder('ブランド画像');
	otherImageFolder = imageFolder.createFolder('素材');

	// 作成したドライブURLを表示させる！
	makeFolderSheet.getRange(2, 8).clearContent();
	let goodsFolderUrl = goodsFolder.getUrl();
	makeFolderSheet.getRange(2, 8).setValue(goodsFolderUrl);
}

//---------------------------------------------------------------------------------------------------------------------
