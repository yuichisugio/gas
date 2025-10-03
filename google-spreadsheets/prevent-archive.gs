// slackチャンネルがアーカイブされないように定期的に送信するための関数
function preventArchivingSlackChannel() {
	let preventArchiveText = '<@xxxxxxxxx>\n「#xxxxxxxxx」のアーカイブ防止用メッセージだよ！';
	sendSlack(preventArchiveText, 'xxxxxxxxx');
}
