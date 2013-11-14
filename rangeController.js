
var rangeController = (function() {

	var lastColumn = sheet.getLastColumn();
	var lastRow = sheet.getLastRow();
	var range = sheet.getRange(1, 1, lastRow, lastColumn);
	var values = range.getValues();
	var backgrounds = range.getBackgrounds();


	

	function getValues() {
		return values;
	}
	function getBackgrounds() {
		return backgrounds;
	}
	function save() {
		sheet
			.getRange(1, 1, values.length, values[0].length)
			.setValues(values)
			.setBackgrounds(backgrounds);
	}

	function getWidth() {
		return values[0].length;
	}
	function getHeight() {
		return values.length;
	}
	function insertColumnsBefore(row, num) {
		sheet.insertColumnsBefore(row, num);
		for (var i = 0 ; i < values.length ; ++i) {
			for (var n = 0 ; n < num ; ++n) {
				values[i].splice(row - 1, 0, null);
				backgrounds[i].splice(row - 1, 0, null);
			}
		}
		return this;
	}
	function insertColumnsAfter(row, num) {
		sheet.insertColumnsAfter(row, num);
		for (var i = 0 ; i < values.length ; ++i) {
			for (var n = 0 ; n < num ; ++n) {
				values[i].splice(row, 0, null);
				backgrounds[i].splice(row, 0, null);
			}
		}
		return this;
	}
	function deleteColumns(row, num) {
		sheet.deleteColumns(row, num);
		for (var i = 0 ; i < values.length ; ++i) {
			values[i].splice(row - 1, num);
			backgrounds[i].splice(row - 1, num);
		}
		return this;
	}

	return {
		getValues: getValues,
		getBackgrounds: getBackgrounds,

		save: save,

		insertColumnsBefore: insertColumnsBefore,
		insertColumnsAfter: insertColumnsAfter,
		deleteColumns: deleteColumns,

		getWidth: getWidth,
		getHeight: getHeight,
	};
})();
