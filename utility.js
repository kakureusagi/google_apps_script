




function toast(title, str, second) {
	spreadsheet.toast(str, title, second ? second: 5);
}

function isNumber(value) {
	return !isNaN(value);
}

function isDate(value) {
	return value instanceof Date;
}

function isUndefined(value) {
	return typeof value === 'undefined';
}

function isNull(value) {
	return value === null;
};



