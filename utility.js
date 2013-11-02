
/**
 * トースト表示
 * @param  {string} str 表示文字列
 * @param  {number} second 表示時間（秒）。デフォルトでは５秒
 */
function toast(str, second) {
	spreadsheet.toast(str, second ? second : 5);
}

