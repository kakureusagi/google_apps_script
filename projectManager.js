
//time
var SECOND = 1000;
var MINITE = SECOND * 60;
var HOUR = MINITE * 60;
var DAY = HOUR * 24;
var YOUBI = ['日', '月', '火', '水', '木', '金', '土'];
var SPECIAL = {
	HOLIDAY: '休',
	NORMAL: '出',
};

var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet = SpreadsheetApp.getActiveSheet();




//position
var AREA_ROOT_X = 5;
var AREA_ROOT_Y = 6;
var SPECIAL_DAY_Y = AREA_ROOT_Y - 4;
var YEAR_Y = AREA_ROOT_Y - 3;
var MONTH_Y = AREA_ROOT_Y - 2;
var DAY_Y = AREA_ROOT_Y - 1;
var PROJECT_START_X = 2;
var PROJECT_START_Y = 1;
var PROJECT_END_X = 2;
var PROJECT_END_Y = 2;
var START_X = 3;
var PERIOD_X = 4;

//color
var COLOR = {
	HOLIDAY: '#ff0000', //土日
	NATIONAL_HOLIDAY: '#aa6666', //祝日
	NONE: '#ffffff', //平日
	UNFINISHED: '#00ff00',
	FINISHED: '#555555',
};

//キャッシュ用の名前
var CACHE = {
	HOLIDAY_TIMES: 'holidayTimes',
};

var FONT_SIZE = 6;
var ALIGN = 'left';




function onOpen(e) {
}

function onEdit(e) {
}



function edit(e) {
	var analyzer = new Analyzer('onEdit');

	try {
		check(e);
	}
	catch(e) {
		Logger.log(e);
	}

	analyzer.set('all time');
}

function check(e) {
	if (e.range.getWidth() != 1 || e.range.getHeight() != 1) {
		return;
	}

	var y = e.range.getRow();
	var x = e.range.getColumn();

	if (project.isUpdate(x, y)) {
		project.updateCache();
		var period = project.getPeriod();
		holiday.update(period.start, period.end);
		project.update();
		row.updateAll();
		return;
	}

	if (row.isUpdate(x, y)) {
		project.update();
		row.updateCache(y);
		row.updateAll();
		return;
	}

	if (project.isSpecialColumn(x, y)) {
		project.update();
		row.updateAll();
		return;
	}
}

function prolongCache() {
	cache.prolong();
}


function updateHoliday() {
	var period = project.getPeriod();
	if (!period) return;

	holiday.update(period.start, period.end);
}




var project = (function() {
	var CACHE_NAME = 'project_cache';

	function getDiffWidth(startTimeInstance, endTimeInstance) {
		return startTimeInstance.diffDay(endTimeInstance) + 1
	}

	function getPeriodFromCell() {
		var startAndEnd = sheet.getRange(PROJECT_START_Y, PROJECT_START_X, 2, 1).getValues();
		var start = startAndEnd[0][0];
		var end = startAndEnd[1][0];

		if (!isDate(start) || !isDate(end)) {
			return null;
		}

		var startInstance = timeController.get(start);
		var endInstance = timeController.get(end);
		if (startInstance.millisecond >= endInstance.millisecond) {
			return null;
		}

		return {
			start: timeController.get(start),
			end: timeController.get(end),
		};
	}

	/**
	 * 特別な列として扱うかどうか
	 * @param  {number} x 
	 * @param  {number} y 
	 * @return {boolean} 
	 */
	function isSpecialColumn(x, y) {
		if (x >= AREA_ROOT_X && y == SPECIAL_DAY_Y) {
			return true;
		}
		return false;
	}

	/**
	 * プロジェクトの期間を更新する必要があるかどうか
	 * @param  {number} x onEditで飛んできた座標
	 * @param  {number} y onEditで飛んできた座標
	 * @return {boolean}
	 */
	function isUpdate(x, y) {
		if ((x == PROJECT_START_X && y == PROJECT_START_Y) || (x == PROJECT_END_X && y == PROJECT_END_Y)) {
			var period = getPeriodFromCell();
			if (!period) {
				return false;
			}

			return true;
		}

		return false;
	}

	/**
	 * プロジェクトの期間を{start: TimeInstance, end: TimeInstance}という形で取得
	 * @return {Objedt}
	 */
	function getPeriod() {
		if (!cache.is(CACHE_NAME)) return null;

		var temp = cache.get(CACHE_NAME);
		return {
			start: timeController.get(temp.start),
			end: timeController.get(temp.end),
		};
	}

	/**
	 * プロジェクトの期間の幅を取得
	 * @return {number} 
	 */
	function getWidth() {
		if (!cache.is(CACHE_NAME)) return 0;

		var item = cache.get(CACHE_NAME);
		return item.width;
	}

	/**
	 * プロジェクトの期間の高さを取得
	 * @return {number} 
	 */
	function getHeight() {
		return sheet.getLastRow() - (AREA_ROOT_Y - 1);
	}

	/**
	 * 現在セルとして設定されているプロジェクトの期間を取得
	 * @return {Object} 
	 */
	function getCurrentPeriod() {
		function getYMD(x) {
			var range = sheet.getRange(YEAR_Y, x, 3, 1).getValues();
			var year = range[0][0];
			var month = range[1][0];
			var day = range[2][0];
			if (!year || !month || !day) {
				return null;
			}

			return timeController.get(year, month, day);
		}

		var temp = {
			start: getYMD(AREA_ROOT_X),
			end: getYMD(sheet.getLastColumn()),
		};
		if (!temp.start || !temp.end) return null;

		return temp;
	}

	/**
	 * プロジェクトの期間を更新する
	 */
	function updateCache() {
		var period = getPeriodFromCell();

		var temp = {
			start: period.start.millisecond,
			end: period.end.millisecond,
			width: getDiffWidth(period.start, period.end),
		};

		cache.set(CACHE_NAME, temp);
	}

	/**
	 * プロジェクトのマスを更新する
	 */
	function update() {
		var analyzer = new Analyzer('updateProjectPeriod');

		var projectPeriod = project.getPeriod();
		if (!projectPeriod) return;

		var currentPeriod = getCurrentPeriod();
		analyzer.set('get project and current time.');

		if (currentPeriod) {
			//既に入力済みなので、追加・修正する
			var startDiff = projectPeriod.start.diffDay(currentPeriod.start);
			if (startDiff > 0) {
				//入力済みの日にちと現在の設定値を比較して、現在の方が開始日が前なら追加、後なら削除
				if (projectPeriod.start.millisecond < currentPeriod.start.millisecond) {
					sheet.insertColumnsBefore(AREA_ROOT_X, startDiff);
					analyzer.set('insert columns').back();
				}
				else if (currentPeriod.start.millisecond < projectPeriod.start.millisecond) {
					var currentWidth = getDiffWidth(currentPeriod.start, currentPeriod.end);
					if (startDiff < currentWidth) {
						sheet.deleteColumns(AREA_ROOT_X, startDiff);
						analyzer.set('delete different columns').back();
					}
				}
			}
		}

		var width = getWidth();
		var lastColumn = sheet.getLastColumn();
		if (lastColumn >= AREA_ROOT_X + width) {
			//不要な領域を削除
			sheet.deleteColumns(AREA_ROOT_X + width, lastColumn - (AREA_ROOT_X - 1 + width));
			analyzer.set('delete unnecessary columns').back();
		}

		/**
		 * 日付を入力
		 */

		//必要な日付を予め計算
		var times = [];
		for (var x = 0 ; x < width ; ++x) {
			times[x] = timeController.instance(projectPeriod.start).plusDay(x).get();
		}

		analyzer.set('compare project time width current time');

		var height = getHeight();

		function initializeAll() {
			sheet
				.getRange(YEAR_Y, AREA_ROOT_X, 3, width)
				.setFontSize(FONT_SIZE)
				.setHorizontalAlignment(ALIGN);
		}
		function initializeYearAndMonth(y, yearOrMonth) {
			var current = 0;
			function mergeAndSet(x, value) {
				sheet
					.getRange(y, AREA_ROOT_X + current, 1, x - current + 1)
					.setValue(value);
				current = x + 1;
			}

			var startValue = times[0][yearOrMonth];
			for (var x = 1 ; x < width ; ++x) {
				if (times[x][yearOrMonth] != startValue) {
					mergeAndSet(x - 1, startValue);
					startValue = times[x][yearOrMonth];
				}

				if (x == width - 1) {
					mergeAndSet(x, times[x][yearOrMonth]);
				}
			}
		}
		function initializeDay() {
			for (var x = 0 ; x < width ; ++x) {
				sheet
					.autoResizeColumn(AREA_ROOT_X + x)
					.getRange(DAY_Y, AREA_ROOT_X + x)
					.setValue(times[x].day);
			}
		}
		initializeAll();
		initializeYearAndMonth(YEAR_Y, 'year');
		initializeYearAndMonth(MONTH_Y, 'month');
		initializeDay();
		
		analyzer.set('initialize year, month and day');

		/**
		 * 日付に色づけ
		 */

		var specialColumn = sheet.getRange(SPECIAL_DAY_Y, AREA_ROOT_X, 1, width).getValues()[0];

		var targetColorRange = null;
		var time = null;
		var isHoliday = false;
		var holidayTimes = holiday.getTimes();
		for (var x = 0 ; x < width ; ++x) {

			targetColorRange = sheet.getRange(AREA_ROOT_Y, AREA_ROOT_X + x, height, 1)
			time = times[x];

			//休みとして指定されている
			if (specialColumn[x] == SPECIAL.HOLIDAY) {
				targetColorRange.setBackground(COLOR.HOLIDAY);
				continue;
			}

			//出勤日として指定されている
			if (specialColumn[x] == SPECIAL.NORMAL) {
				targetColorRange.setBackground(COLOR.NONE);
				continue;
			}

			//祝日
			isHoliday = false;
			for (var i = 0 ; i < holidayTimes.length ; ++i) {
				if (time.millisecond == holidayTimes[i].millisecond) {
					targetColorRange.setBackground(COLOR.NATIONAL_HOLIDAY);
					isHoliday = true;
					break;
				}
			}
			if (isHoliday) {
				continue;
			}

			//土日
			if (time.youbi == '土' || time.youbi == '日') {
				targetColorRange.setBackground(COLOR.HOLIDAY);
				continue;
			}

			//平日
			targetColorRange.setBackground(COLOR.NONE);
		}

		analyzer.set('set color on days');
	}


	return {
		isSpecialColumn: isSpecialColumn,

		isUpdate: isUpdate,
		updateCache: updateCache,
		update: update,

		getPeriod: getPeriod,

		getWidth: getWidth,
		getHeight: getHeight,
	};
})();

var row = (function() {

	var CACHE_NAME = 'row_cache';



	function getStartAndTimeFromCell(y) {
		var startAndTime = sheet.getRange(y, START_X, 1, 2).getValues();
		var start = startAndTime[0][0];
		var time = Math.floor(startAndTime[0][1]);

		if (!isDate(start) || !isNumber(time) || time <= 0) {
			return null;
		}

		return {
			start: timeController.get(start),
			time: time,
		};
	}

	function getStartAndTime(y) {
		if (!cache.isRow(y)) return null;

		var temp = cache.getRow(y, CACHE_NAME);
		return {
			start: timeController.get(temp.start),
			time: temp.time,
		};
	}

	function updateCache(y) {
		var startAndTime = getStartAndTimeFromCell(y);
		var temp = {
			start: startAndTime.start.millisecond,
			time: startAndTime.time,
		};
		cache.setRow(y, CACHE_NAME, temp);
	}

	function isUpdate(x, y) {
		if (y < AREA_ROOT_Y) return false;
		if (x != START_X && x != PERIOD_X) return false;

		var temp = getStartAndTimeFromCell(y);
		if (!temp) {
			return false;
		}

		return true;
	}

	function update(y) {
		var projectPeriod = project.getPeriod();
		if (!projectPeriod) return;

		var temp = getStartAndTimeFromCell(y);
		if (!temp) return;

		var start = temp.start;
		var time = temp.time;

		//プロジェクト期間より後ろの日付は無視
		if (projectPeriod.end.millisecond < start.millisecond) return;

		if (start.millisecond < projectPeriod.start.millisecond) {
			utility.toast('項目の開始日がプロジェクトの開始日よりも前になっています。', y + '行目');
			return;
		}
		

		var width = project.getWidth();
		var backgoundColors = sheet.getRange(AREA_ROOT_Y, AREA_ROOT_X, 1, width).getBackgrounds()[0];
		var count = 0;
		var finishedCount = 0;
		while (finishedCount != time) {
			var t = timeController.instance(start).plusDay(count).get();

			//プロジェクトの最終日を超えてしまった
			if (projectPeriod.end.millisecond < t.millisecond) {
				break;
			}

			//まだプロジェクトの開始前
			if (t.millisecond < projectPeriod.start.millisecond) {
				++count;
				continue;
			}

			//休みはカウントしない
			var diff = t.diffDay(projectPeriod.start);
			if (isHolidayColor(backgoundColors[diff])) {
				++count;
				continue;
			}

			//平日
			sheet.getRange(y, AREA_ROOT_X + diff).setBackground(COLOR.UNFINISHED);

			++count;
			++finishedCount;
		}
	}

	function updateAll() {
		var analyzer = new Analyzer('updateAllRows');

		var height = project.getHeight();
		for (var y = 0 ; y < height ; ++y) {
			if (!isUpdate(PERIOD_X, AREA_ROOT_Y + y)) continue;
			update(AREA_ROOT_Y + y);
		}

		analyzer.set('update all rows');
	}

	return {
		isUpdate: isUpdate,
		update: update,
		updateCache: updateCache,
		updateAll: updateAll,
	};
})();





function isHolidayColor(value) {
	return value == COLOR.HOLIDAY || value == COLOR.NATIONAL_HOLIDAY;
}



