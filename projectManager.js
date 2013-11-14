
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

//position
var AREA_ROOT_X = 6;
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
var PROGRESS_X = 5;

//color
var COLOR = {
	TODAY: '#fcff84',

	HOLIDAY: '#ff0000', //土日
	NATIONAL_HOLIDAY: '#ff8888', //祝日
	NONE: '#ffffff', //平日
	UNFINISHED: '#d5e5ff',
	FINISHED: '#1660f4',
};

//キャッシュ用の名前
var CACHE = {
	HOLIDAY_TIMES: 'holidayTimes',
};

var FONT_SIZE = 6;
var ALIGN = 'left';



// global
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet = SpreadsheetApp.getActiveSheet();




/**
 * for simple trigger
 */
function onOpen(e) {
	spreadsheet.addMenu('スケジュール', [
		{name: '初期化', functionName: 'initializeTrigger'},
		{name: '更新', functionName: 'updateAll'},
	]);
}


/**
 * for custom trgger
 */

function open(e) {
	project.setToday();
	rangeController.save();
}

function edit(e) {
	var analyzer = new Analyzer('edit');

	try {
		check(e);
	}
	catch(e) {
		Logger.log(e);
		toast('Error!', e);
	}

	analyzer.set('all time');
}

function checkToday() {
	project.setToday();
	rangeController.save();
}

function prolongCache() {
	cache.prolong();
}

function initializeTrigger() {
	trigger.initialize();
}

function updateAll() {
	project.updateHoliday();
	project.update();
	row.update();

	rangeController.save();
}





function check(e) {
	var analyzer = new Analyzer('check');
	var left = e.range.getColumn();
	var top = e.range.getRow();
	var values = e.range.getValues();

	var projectChecker = false;
	var rowChecker = {};

	if (!lock.start()) {
		toast('競合エラー！', '処理が追いついていません。少し待ってから操作して下さいね。');
	}

	for (var x = left ; x < left + values[0].length ; ++x) {
		for (var y = top ; y < top + values.length ; ++y) {

			if (project.isUpdate(x, y)) {
				project.updateHoliday();
				analyzer.set('holiday update');
				project.update();
				analyzer.set('project update');
				row.update();
				analyzer.set('row update');
				rangeController.save();
				analyzer.set('range save');
				return;
			}

			if (row.isUpdate(x, y) || prject.isSpecialColumn(x, y)) {
				project.update();
				analyzer.set('project update');
				row.update();
				analyzer.set('row update');
				rangeController.save();
				analyzer.set('range save');
				return;
			}
		}
	}
}




/**
 * トリガー操作モジュール
 */

var trigger = (function() {

	function initialize() {
		//一旦全てのトリガを削除
		var triggers = ScriptApp.getProjectTriggers();
		triggers.forEach(function(t, i) {
			ScriptApp.deleteTrigger(t);
		});

		ScriptApp.newTrigger('open')
			.forSpreadsheet(spreadsheet.getId())
			.onOpen()
			.create();

		ScriptApp.newTrigger('edit')
			.forSpreadsheet(spreadsheet.getId())
			.onEdit()
			.create();
			
		ScriptApp.newTrigger('prolongCache')
			.timeBased()
			.everyHours(1)
			.create();
			
		ScriptApp.newTrigger('checkToday')
			.timeBased()
			.everyHours(1)
			.create();
	}

	return {
		initialize: initialize,
	};
})();



var project = (function() {
	var CACHE_NAME = 'project_cache';

	function getDiffWidth(startTimeInstance, endTimeInstance) {
		return startTimeInstance.diffDay(endTimeInstance) + 1;
	}

	function getPeriodFromCell() {
		var values = rangeController.getValues();
		var start = values[PROJECT_START_Y - 1][PROJECT_START_X - 1];
		var end = values[PROJECT_START_Y][PROJECT_START_X - 1];
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
		return getPeriodFromCell();
	}

	/**
	 * プロジェクトの期間の幅を取得
	 * @return {number} 
	 */
	function getWidth() {
		var p = getPeriod();
		return getDiffWidth(p.start, p.end);
	}

	/**
	 * プロジェクトの期間の高さを取得
	 * @return {number} 
	 */
	function getHeight() {
		return rangeController.getValues().length - (AREA_ROOT_Y - 1);
	}

	function getCurrentWidth() {
		return rangeController.getWidth() - (AREA_ROOT_X - 1);
	}

	function getCurrentHeight() {
		//todo:
	}

	/**
	 * 現在セルとして設定されているプロジェクトの期間を取得
	 * @return {Object} 
	 */
	function getCurrentPeriod() {
		var values = rangeController.getValues();
		function getYMD(x) {
			var year = values[YEAR_Y - 1][x - 1];
			var month = values[YEAR_Y][x - 1];
			var day = values[YEAR_Y + 1][x - 1];
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

	function updateHoliday() {
		var period = project.getPeriod();
		if (!period) return;

		holiday.update(period.start, period.end);
	}

	/**
	 * プロジェクトのマスを更新する
	 */
	function update() {
		var projectPeriod = project.getPeriod();
		if (!projectPeriod) return;

		/**
		 * 現在の日付と比較して、列の追加・削除が必要なら行う
		 */
		function comparePeriod() {
			var currentPeriod = getCurrentPeriod();
			if (currentPeriod) {
				//既に入力済みなので、追加・修正する
				var startDiff = projectPeriod.start.diffDay(currentPeriod.start);
				if (startDiff > 0) {
					//入力済みの日にちと現在の設定値を比較して、現在の方が開始日が前なら追加、後なら削除
					if (projectPeriod.start.millisecond < currentPeriod.start.millisecond) {
						rangeController.insertColumnsBefore(AREA_ROOT_X, startDiff);
					}
					else if (currentPeriod.start.millisecond < projectPeriod.start.millisecond) {
						var currentWidth = getDiffWidth(currentPeriod.start, currentPeriod.end);
						if (startDiff < currentWidth) {
							rangeController.deleteColumns(AREA_ROOT_X, startDiff);
						}
					}
				}
			}
		}
		comparePeriod();
		

		/**
		 * プロジェクトの期間と列の長さを合わせる
		 */
		function adjustProjectWidth() {
		 	var currentWidth = getCurrentWidth();
		 	var projectWidth = getWidth();
		 	if (currentWidth > projectWidth) {
		 		rangeController.deleteColumns(AREA_ROOT_X + projectWidth - 1, currentWidth - projectWidth);
		 	}
		 	else if (currentWidth < projectWidth) {
		 		rangeController.insertColumnsAfter(AREA_ROOT_X + currentWidth - 1, projectWidth - currentWidth);
		 	}
		}
		adjustProjectWidth();

		/**
		 * 日付を入力
		 */

		//必要な日付を予め計算
		var width = getWidth();
		var height = getHeight();
		var times = [];
		for (var x = 0 ; x < width ; ++x) {
			times[x] = timeController.instance(projectPeriod.start).plusDay(x).get();
		}

		function inputDate() {
			var values = rangeController.getValues();
			for (var x = 0 ; x < width ; ++x) {
				values[YEAR_Y - 1][AREA_ROOT_X - 1 + x] = times[x].year;
				values[MONTH_Y - 1][AREA_ROOT_X - 1 + x] = times[x].month;
				values[DAY_Y - 1][AREA_ROOT_X - 1 + x] = times[x].day;

				sheet.setColumnWidth(AREA_ROOT_X + x, 20);
			}

			sheet
				.getRange(YEAR_Y, AREA_ROOT_X, 3, width)
				.setHorizontalAlignment(ALIGN)
				.setFontSize(FONT_SIZE);
		}
		inputDate();

		/**
		 * 日付に色づけ
		 */
		function makeColor() {
			var values = rangeController.getValues();
			var backgrounds = rangeController.getBackgrounds();
			var specialColumn = values[SPECIAL_DAY_Y - 1];

			var time = null;
			var isHoliday = false;
			var holidayTimes = holiday.getTimes();

			function isHolidayTime(time) {
				for (var i = 0 ; i < holidayTimes.length ; ++i) {
					if (time.millisecond == holidayTimes[i].millisecond) {
						return true;
					}
				}
				return false;
			}
			for (var x = 0 ; x < width ; ++x) {
				time = times[x];
				var color = "";
				if (specialColumn[AREA_ROOT_X - 1 + x] == SPECIAL.HOLIDAY) {
					color = COLOR.HOLIDAY;
				}
				else if (specialColumn[AREA_ROOT_X - 1 + x] == SPECIAL.NORMAL) {
					color = COLOR.NONE;
				}
				else if (isHolidayTime(time)) {
					color = COLOR.NATIONAL_HOLIDAY;
				}
				else if (time.youbi == '土' || time.youbi == '日') {
					color = COLOR.HOLIDAY;
				}
				else {
					color = COLOR.NONE;
				}

				for (var y = 0 ; y < height ; ++y) {
					backgrounds[AREA_ROOT_Y - 1 + y][AREA_ROOT_X - 1 + x] = color;
				}
			}
		}
		makeColor();
	}

	function setToday() {
		var period = getPeriod();
		if (!period) return;

		//一旦全範囲をリセット
		var backgrounds = rangeController.getBackgrounds();
		for (var y = 0 ; y < 3 ; ++y) {
			for (var x = 0 ; x < getWidth() ; ++x) {
				backgrounds[YEAR_Y - 1 + y][AREA_ROOT_X - 1 + x] = COLOR.NONE;
			}
		}

		var today = timeController.get();
		if (today.millisecond < period.start.millisecond || period.end.millisecond < today.millisecond) return;

		var diff = today.diffDay(period.start);
		for (y = 0 ; y < 3 ; ++y) {
			backgrounds[YEAR_Y - 1 + y][AREA_ROOT_X - 1 + diff] = COLOR.TODAY;
		}
	}


	return {
		isSpecialColumn: isSpecialColumn,
		isUpdate: isUpdate,

		updateHoliday: updateHoliday,
		update: update,

		getPeriod: getPeriod,

		getWidth: getWidth,
		getHeight: getHeight,

		setToday: setToday,
	};
})();

var row = (function() {

	var CACHE_NAME = 'row_cache';



	function validateInfo(start, time, progress) {
		if (!isDate(start) || !isNumber(time) || time <= 0) {
			return false;
		}
		return true;
	}

	function getRowInfoFromCell(row) {
		var values = rangeController.getValues();
		var start = values[row - 1][START_X - 1];
		var time = values[row - 1][START_X];
		var progress = values[row - 1][START_X + 1];
		progress = progress ? progress : 0;

		if (!validateInfo(start, time, progress)) {
			return null;
		}

		return {
			start: timeController.get(start),
			time: time,
			progress: progress,
		};
	}

	function getRowInfo(y) {
		return getRowInfoFromCell(y);
	}

	function updateCache(y) {
		var info = getRowInfoFromCell(y);
		if (!info) return;

		var temp = {
			start: info.start.millisecond,
			time: info.time,
			progress: info.progress
		};
		cache.setRow(y, CACHE_NAME, temp);
	}

	function updateCacheAll() {
		var info = sheet.getRange(AREA_ROOT_Y, START_X, project.getHeight(), 3).getValues();
		for (var i = 0 ; i < info.length ; ++i) {
			var start = info[i][0];
			var time = Math.floor(info[i][1]);
			var progress = info[i][2] ? info[i][2] : 0;
			var data = null;
			if (validateInfo(start, time, progress)) {
				data = {
					start: timeController.get(start).millisecond,
					time: time,
					progress:progress,
				};
			}

			cache.setRow(AREA_ROOT_Y + i, CACHE_NAME, data);
		}
	}

	function isUpdate(x, y) {
		if (y < AREA_ROOT_Y) return false;
		if (x != START_X && x != PERIOD_X && x != PROGRESS_X) return false;

		var info = getRowInfoFromCell(y);
		if (!info) {
			return false;
		}

		return true;
	}

	function update(y) {
		var projectPeriod = project.getPeriod();
		if (!projectPeriod) return;

		var infos = [];
		if (y) {
			infos.push(getRowInfo(y));
		}
		else {
			for (var h = 0 ; h < project.getHeight() ; ++h) {
				infos.push(getRowInfo(AREA_ROOT_Y + h));
			}
		}

		var backgrounds = rangeController.getBackgrounds();
		infos.forEach(function(info, index) {
			if (!info) return;

			var start = info.start;
			var time = info.time;
			var progress = info.progress;
			//var posY = y ? y : AREA_ROOT_Y + index;
			var posY = AREA_ROOT_Y + index;

			//プロジェクト期間より後ろの日付は無視
			if (projectPeriod.end.millisecond < start.millisecond) return;

			//プロジェクトの期間より前の日付
			if (start.millisecond < projectPeriod.start.millisecond) {
				toast('期間エラー！', '項目の開始日がプロジェクトの開始日よりも前になっています。' + posY + '行目');
				return;
			}
			
			/**
			 * 色付け
			 */
			
			var count = 0;
			var finishedCount = 0;
			var finishedPos = Math.floor(time * (progress / 100));
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
				if (isHolidayColor(backgrounds[posY - 1][AREA_ROOT_X - 1 + diff])) {
					++count;
					continue;
				}

				//平日
				++finishedCount;
				if (finishedCount > finishedPos) {
					backgrounds[posY - 1][AREA_ROOT_X - 1 + diff] = COLOR.UNFINISHED;
				}
				else {
					backgrounds[posY - 1][AREA_ROOT_X - 1 + diff] = COLOR.FINISHED;
				}

				++count;
			}
		});
	}

	return {
		isUpdate: isUpdate,
		update: update,
		updateCache: updateCache,
		updateCacheAll: updateCacheAll,
	};
})();





function isHolidayColor(value) {
	return value == COLOR.HOLIDAY || value == COLOR.NATIONAL_HOLIDAY;
}



