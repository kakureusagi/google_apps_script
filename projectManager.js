//time
var SECOND = 1000;
var MINITE = SECOND * 60;
var HOUR = MINITE * 60;
var DAY = HOUR * 24;
var YOUBI = ['日', '月', '火', '水', '木', '金', '土'];
var YASUMI = '休';
var SYUKU = '祝';

//position
var AREA_ROOT_X = 5;
var AREA_ROOT_Y = 6;
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


//global
var sheet = SpreadsheetApp.getActiveSheet();
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();






function onOpen(e) {
	updateHolidayTimes();
}

function onEdit(e) {
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

	var analyzer = new Analyzer('check');

	//
	if (isUpdateProjectPeriod(x, y)) {
		updateHolidayTimes();
		updateProjectPeriod();
		updateAllRows();
		return;
	}

	if (isUpdateRow(x, y)) {
		updateProjectPeriod();
		updateRow(y);
		return;
	}
}





var project = (function() {
	return {
		updatePeriodCache: function() {

		},

		updatePeriod: function() {

		},

		isUpatePeriod: function(x, y) {

		},

		getPeriod: function() {

		},

		getWidth: function() {

		},

		getHeight: function() {

		},
	};
})();

var rowController = (function() {
	return {
	};
})();





/**
 * プロジェクトの開始日・終了日を取得
 * 開始日や終了日が妥当でない場合には、nullが返る
 * @return {Object} {startTime: timeInstance, endTime: timeInstance}
 */
function getProjectTime() {
	var start = sheet.getRange(PROJECT_START_Y, PROJECT_START_X).getValue();
	var end = sheet.getRange(PROJECT_END_Y, PROJECT_END_X).getValue();

	//値がない
	if (!start || !end) {
		return null;
	}

	var startTime = timeController.instance(start).get();
	var endTime = timeController.instance(end).get();

	//開始日・終了日がおかしい
	if (startTime.millisecond > endTime.millisecond) {
		return null;
	}

	return {
		start: startTime,
		end: endTime,
	};
}

/**
 * 現在設定されている日付を取得する
 * 開始日や終了日が妥当でない場合には、nullが返る
 * @return {Object} {start: timeInstance, end: timeInstance}
 */
function getCurrentTime() {
	function getYMD(x) {
		var range = sheet.getRange(YEAR_Y, x, 3, 1).getValues();
		var year = range[0][0];
		var month = range[0][1];
		var day = range[0][2];
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

function updateHolidayAndPeriod() {
	updateHolidayTimes();
	updateProjectPeriod();
}

/**
 * プロジェクトの期間を変更する必要があるかどうか
 * @param  {number} x 
 * @param  {number} y 
 * @return {boolean} 
 */
function isUpdateProjectPeriod(x, y) {
	return (x == PROJECT_START_X && y == PROJECT_START_Y) || (x == PROJECT_END_X && y == PROJECT_END_Y);
}

/**
 * プロジェクトの期間を更新する
 */
function updateProjectPeriod() {
	var analyzer = new Analyzer('updateProjectPeriod');
	
	var projectTime = getProjectTime();
	if (!projectTime) return;

	var currentTime = getCurrentTime();

	analyzer.set('get project and current time.');

	if (currentTime) {
		//既に入力済みなので、追加・修正する
		var startDiff = projectTime.start.diffDay(currentTime.start);
		if (startDiff > 0) {
			//入力済みの日にちと現在の設定値を比較して、現在の方が開始日が前なら追加、後なら削除
			if (projectTime.start.millisecond < currentTime.start.millisecond) {
				sheet.insertColumnsBefore(AREA_ROOT_X, startDiff);
				analyzer.set('insert columns').back();
			}
			else if (currentTime.start.millisecond < projectTime.start.millisecond) {
				var currentWidth = currentTime.start.diffDay(currentTime.end) + 1;
				if (startDiff < currentWidth) {
					sheet.deleteColumns(AREA_ROOT_X, startDiff);
					analyzer.set('delete different columns').back();
				}
			}
		}
	}

	//不要な領域を削除
	var width = getWidth();
	if (sheet.getLastColumn() >= AREA_ROOT_X + width) {
		sheet.deleteColumns(AREA_ROOT_X + width, sheet.getLastColumn() - (AREA_ROOT_X - 1 + width));
		analyzer.set('delete unnecesary columns').back();
	}

	/**
	 * 日付を入力
	 */

	//必要な日付を予め計算
	var times = [0];
	for (var x = 1 ; x <= width ; ++x) {
		times[x] = timeController.instance(projectTime.start).plusDay(x - 1).get();
	}

	//現在値が設定されている値を取得（設定済みのところは無視する）
	var dayValues = sheet.getRange(DAY_Y, AREA_ROOT_X, 1, width).getValues()[0];

	analyzer.set('compare project time width current time');


	function initializeYearAndMonth(y, yearOrMonth) {
		var current = 1;
		function mergeAndSet(x, value) {
			sheet
				.getRange(y, AREA_ROOT_X + current - 1, 1, x - current + 1)
				//.merge()
				.setValue(value)
				.setFontSize(8)
				.setHorizontalAlignment('left');
			current = x + 1;
		}

		var startValue = times[1][yearOrMonth];
		for (var x = 2 ; x <= width ; ++x) {
			if (times[x][yearOrMonth] != startValue) {
				mergeAndSet(x - 1, startValue);
				startValue = times[x][yearOrMonth];
			}

			if (x == width) {
				mergeAndSet(x, times[x][yearOrMonth]);
			}
		}
	}
	function initializeDay() {
		for (var x = 1 ; x <= width ; ++x) {
			sheet
				.autoResizeColumn(AREA_ROOT_X + x - 1)
				.getRange(DAY_Y, AREA_ROOT_X + x - 1)
				.setValue(times[x].day)
				.setFontSize(8);
		}
	}
	initializeYearAndMonth(YEAR_Y, 'year');
	initializeYearAndMonth(MONTH_Y, 'month');
	initializeDay();
	
	analyzer.set('initialize year, month and day');

	/**
	 * 日付に色づけ
	 */

	var height = getHeight();
	var area = sheet.getRange(AREA_ROOT_Y, AREA_ROOT_X, height, width);

	var targetColorRange = null;
	var time = null;
	var isHoliday = false;
	var holidayTimes = getHolidayTimes();
	for (var x = 1 ; x <= width ; ++x) {

		targetColorRange = sheet.getRange(AREA_ROOT_Y, AREA_ROOT_X + x - 1, height, 1)
		time = times[x];

		//祝日
		isHoliday = false;
		for (var i = 0 ; i < holidayTimes.length ; ++i) {
			if (time.millisecond == holidayTimes[i].millisecond) {
				targetColorRange.setBackground(COLOR.NATIONAL_HOLIDAY).setValue(SYUKU);
				isHoliday = true;
				break;
			}
		}
		if (isHoliday) {
			continue;
		}

		//土日
		if (time.youbi == '土' || time.youbi == '日') {
			targetColorRange.setBackground(COLOR.HOLIDAY).setValue(YASUMI);
			continue;
		}

		//既にセットされている平日はそのまま
		if (dayValues[x - 1]) continue;

		//平日
		targetColorRange.setBackground(COLOR.NONE);
	}

	analyzer.set('set color on days');
}

/**
 * 行を修正する必要があるかどうか
 * @param  {number} x 
 * @param  {number} y 
 */
function isUpdateRow(x, y, start, period) {
	if (y < AREA_ROOT_Y) return false;
	if (x != START_X && x != PERIOD_X) return false;

	if (!start) {
		start = sheet.getRange(y, START_X).getValue();
	}
	if (!period) {
		period = sheet.getRange(y, PERIOD_X).getValue();
	}
	period -= 0;

	if (!(start instanceof Date) || isNaN(period) || period < 0) {
		return false;
	}

	return true;
}

/**
 * 行を更新する
 * @param  {[type]} y [description]
 * @return {[type]} [description]
 */
function updateRow(y) {
	var startTime = timeController.get(sheet.getRange(y, START_X).getValue());
	var period = Math.floor(sheet.getRange(y, PERIOD_X).getValue());

	var projectTime = getProjectTime();
	if (!projectTime) return;

	//プロジェクト期間より後ろの日付は無視
	if (projectTime.end.millisecond < startTime.millisecond) return;

	if (startTime.millisecond < projectTime.start.millisecond) {
		toast('項目の開始日がプロジェクトの開始日よりも前になっています。', y + '行目');
		return;
	}
	
	//既に色付けされている部分を初期化
	var rowWidth = sheet.getLastColumn() - START_X + 1;
	var rowBackgrounds = sheet.getRange(y, START_X, 1, rowWidth).getBackgrounds()[0];
	for (var x = 0 ; x < rowWidth ; ++x) {
		if (rowBackgrounds[x] == COLOR.UNFINISHED) {
			sheet.getRange(y, START_X + x).setBackground(COLOR.NONE);
		}
	}

	var count = 0;
	var finishedCount = 0;
	var holidayTimes = getHolidayTimes();
	while (finishedCount != period) {
		var time = timeController.instance(startTime).plusDay(count).get();

		//プロジェクトの最終日を超えてしまった
		if (projectTime.end.millisecond < time.millisecond) {
			break;
		}

		//まだプロジェクトの開始前
		if (time.millisecond < projectTime.start.millisecond) {
			++count;
			continue;
		}

		//土日はカウントしない
		if (time.youbi == '土' || time.youbi == '日') {
			++count;
			continue;
		}

		//祝日はカウントしない
		var isHoliday = false;
		for (var i = 0 ; i < holidayTimes.length ; ++i) {
			if (time.millisecond == holidayTimes[i].millisecond) {
				isHoliday = true;
				break;
			}
		}
		if (isHoliday) {
			++count;
			continue;
		}

		//平日
		var diff = time.diffDay(projectTime.start);
		var cell = sheet.getRange(y, AREA_ROOT_X + diff);
		cell.setBackground(COLOR.UNFINISHED);

		++count;
		++finishedCount;
	}
}

/**
 * それぞれの行を更新
 */
function updateAllRows() {
	var analyzer = new Analyzer('updateAllRows');

	var height = getHeight();
	var items = sheet.getRange(AREA_ROOT_Y, START_X, height, 2).getValues();
	for (var y = 0 ; y < height ; ++y) {
		if (!items[y][0] || !items[y][1]) continue;

		if (!isUpdateRow(PERIOD_X, AREA_ROOT_Y + y, items[y][0], items[y][1])) continue;

		updateRow(AREA_ROOT_Y + y);
	}

	analyzer.set('update all rows');
}


/**
 * 期間として有効な幅を取得する
 * @return {number} 
 */
function getWidth() {
	var projectTime = getProjectTime();
	return projectTime.start.diffDay(projectTime.end) + 1;
}

/**
 * 有効な高さを取得する
 * @return {number} 
 */
function getHeight() {
	return sheet.getLastRow() - (AREA_ROOT_Y - 1);
}








/**
 * 休日情報を取得する
 * @return {Array} 休日のTimeInstanceが格納された配列
 */
function getHolidayTimes() {
	var holidayTimes = cache.get(getCacheName());

	if (!holidayTimes) {
		updateHolidayTimes();
		holidayTimes = cache.get(getCacheName());
	}

	return holidayTimes;
}

/**
 * 休日情報を更新する
 */
function updateHolidayTimes() {
	var projectTime = getProjectTime();
	if (!projectTime) return;

	var holidays = CalendarApp
		.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com")
		.getEvents(
			projectTime.start.date,
			timeController.instance(projectTime.end).plusDay(1).get().date
		);

	var holidayTimes = [];
	for (var i = 0 ; i < holidays.length ; ++i) {
		holidayTimes.push(timeController.instance(holidays[i].getStartTime()).get());
	}

	cache.set(getCacheName(), holidayTimes);
}

/**
 * 祝日のキャッシュ用の名前を生成する
 * キャッシュの名前はシート毎に固有となる
 * @return {string} 
 */
function getCacheName() {
	return spreadsheet.getId() + sheet.getSheetId() + CACHE.HOLIDAY_TIMES;
}






