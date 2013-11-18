
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
	UNFINISHED: '#c5d5ef',
	FINISHED: '#1660f4',
};

var FONT_SIZE = 4;
var CELL_SIZE = 20;
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
		{name: '開始日でソート', functionName: 'sortByStart'},
		{name: '担当者でソート', functionName: 'sortByStaff'},
	]);

	//open(e);
}


/**
 * for custom trgger
 */

/**
 * オープン時のトリガ
 */
function open(e) {
	tryLock(function() {
		project.updateCache();
		updateHolidayCache();
		row.updateCacheAll();
		project.setToday();
	});
}

/**
 * 編集時のトリガ
 */
function edit(e) {
	var analyzer = new Analyzer('edit');

	try {
		tryLock(function() {
			check(e);
		});
	}
	catch(e) {
		Logger.log(e);
		toast(e);
	}

	analyzer.set('all time');
}

/**
 * 時間で呼ばれるトリガ。今日の日付をチェック
 */
function checkToday() {
	tryLock(function() {
		project.setToday();
	});
}

/**
 * 時間で呼ばれるトリガー。キャシュの生存期間を伸ばす
 */
function prolongCache() {
	tryLock(function() {
		cache.prolong();
	});
}


/**
 * ツールバー系
 */

/**
 * トリガーを初期化する
 */
function initializeTrigger() {
	trigger.initialize();
}

/**
 * 全てのデータを更新する
 */
function updateAll() {
	tryLock(function() {
		var analyzer = new Analyzer('updateAll');

		project.updateCache();
		analyzer.set('update project cache');
		updateHolidayCache();
		analyzer.set('update holiday cache');
		row.updateCacheAll();
		analyzer.set('update rows cache');

		project.update();
		analyzer.set('update project');
		row.update();
		analyzer.set('update rows');
		project.setToday();
		analyzer.set('update today color');
	});
}

/**
 * タスクの開始日でソート
 */
function sortByStart() {
	tryLock(function() {
		sort.start();
	});
}

/**
 * 担当者とタスクの開始日でそ～t
 */
function sortByStaff() {
	tryLock(function() {
		sort.staff();
	});
}





/**
 * セル編集時の処理
 */
function check(e) {
	var analyzer = new Analyzer('check');

	var left = e.range.getColumn();
	var top = e.range.getRow();
	var values = e.range.getValues();
	var rowChecker = {};
	analyzer.set('check event');

	for (var x = left ; x < left + values[0].length ; ++x) {
		for (var y = top ; y < top + values.length ; ++y) {

			if (project.isUpdate(x, y)) {
				project.updateCache();
				analyzer.set('update project cache');
				updateHolidayCache();
				analyzer.set('update holiday cache');
				row.updateCacheAll();
				analyzer.set('update rows cache');

				project.update();
				analyzer.set('update project')
				row.update();
				analyzer.set('update rows');
				project.setToday();
				analyzer.set('update today color');
				return;
			}

			if (row.isUpdate(x, y)) {
				if (rowChecker[y]) continue;
				rowChecker[y] = true;

				project.update(y);
				analyzer.set('update project one row')
				row.updateCache(y);
				analyzer.set('update row cache');
				row.update(y);
				analyzer.set('update row');
				continue;
			}

			if (project.isSpecialColumn(x, y)) {
				project.update();
				analyzer.set('update project')
				row.update();
				analyzer.set('update rows');
				return;
			}
		}
	}
	analyzer.set('all time');
}








/**
 * トリガー操作モジュール
 */

var trigger = (function() {

	/**
	 * トリガーを初期化する
	 */
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




/**
 * プロジェクトの期間操作モジュール
 */
var project = (function() {
	var CACHE_NAME = 'project_cache';

	/**
	 * 特定の2地点でセルのセルの幅を取得
	 * @param  {TimeInstance} startTimeInstance 
	 * @param  {TimeInstance} endTimeInstance 
	 * @return {number} セルの幅
	 */
	function getDiffWidth(startTimeInstance, endTimeInstance) {
		return startTimeInstance.diffDay(endTimeInstance) + 1
	}

	/**
	 * プロジェクトの初日と最終日と幅を取得
	 * @return {Object} {start: timeInstance, end: timeInstance, width:number}
	 */
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
			width: getDiffWidth(startInstance, endInstance)
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
	 * プロジェクトの期間の初日と最終日と幅を取得
	 * @return {Objedt}
	 */
	function getPeriod() {
		if (!cache.is(CACHE_NAME)) return null;

		var temp = cache.get(CACHE_NAME);
		return {
			start: timeController.get(temp.start),
			end: timeController.get(temp.end),
			width: temp.width,
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

		temp.width = getDiffWidth(temp.start, temp.end);
		return temp;
	}

	/**
	 * プロジェクトの期間を更新する
	 */
	function updateCache() {
		var period = getPeriodFromCell();
		if (!period) return;

		var temp = {
			start: period.start.millisecond,
			end: period.end.millisecond,
			width: period.width,
		};

		cache.set(CACHE_NAME, temp);
	}

	/**
	 * プロジェクトのマスを更新する
	 * @param {number} y 指定がなければプロジェクト全体を、指定があれば特定の行だけ更新
	 */
	function update(y) {
		var projectPeriod = project.getPeriod();
		if (!projectPeriod) return;

		/**
		 * 列の幅をプロジェクトの期間に合わせる
		 */
		function adjustProjectWidth() {
			var width = projectPeriod.width;
			var currentMaxWidth = sheet.getMaxColumns() - (AREA_ROOT_X - 1);
			if (width < currentMaxWidth) {
				sheet.deleteColumns(AREA_ROOT_X + width - 1, currentMaxWidth - width);
			}
			else if (currentMaxWidth < width) {
				sheet.insertColumnsAfter(AREA_ROOT_X + currentMaxWidth - 1, width - currentMaxWidth);
			}
		}
		adjustProjectWidth();


		var width = getWidth();
		var height = getHeight();

		//必要な日付を予め計算
		var times = [];
		for (var x = 0 ; x < width ; ++x) {
			times[x] = timeController.instance(projectPeriod.start).plusDay(x).get();
		}

		/**
		 * 日付を入力
		 */
		function inputDate() {
			var range = sheet.getRange(YEAR_Y, AREA_ROOT_X, 3, width);
			var values = range.getValues();
			for (var x = 0 ; x < width ; ++x) {
				values[0][x] = times[x].year;
				values[1][x] = times[x].month;
				values[2][x] = times[x].day;

				sheet.setColumnWidth(AREA_ROOT_X + x, CELL_SIZE);
			}

			range
				.setValues(values)
				.setHorizontalAlignment(ALIGN)
				.setFontSize(FONT_SIZE);
		}

		//日付の設定は全体の変更時だけで十分
		if (!y) {
			inputDate();
		}

		/**
		 * 日付に色づけ
		 */
		function inputColor() {
			var specialColumn = sheet.getRange(SPECIAL_DAY_Y, AREA_ROOT_X, 1, width).getValues()[0];

			var targetColorRange = null;
			var time = null;
			var holidayTimes = holiday.getTimes();

			function isHoliday(time) {
				for (var i = 0 ; i < holidayTimes.length ; ++i) {
					if (time.millisecond == holidayTimes[i].millisecond) {
						return true;
					}
				}
				return false;
			}

			//１列ずつ処理
			for (var x = 0 ; x < width ; ++x) {

				/*
				 * 色付け範囲決定
				 */
				if (y) {
					targetColorRange = sheet.getRange(y, AREA_ROOT_X + x);
				}
				else {
					targetColorRange = sheet.getRange(AREA_ROOT_Y, AREA_ROOT_X + x, height, 1);
				}
				time = times[x];

				/*
				 * 色決定
				 */
				var color = COLOR.NONE; //平日
				if (specialColumn[x] == SPECIAL.HOLIDAY) {
					//休みとして指定されている
					color = COLOR.HOLIDAY;
				}
				else if (specialColumn[x] == SPECIAL.NORMAL) {
					//出勤日として指定されている
					color = COLOR.NONE;
				}
				else if (isHoliday(time)) {
					//祝日
					color = COLOR.NATIONAL_HOLIDAY;
				}
				else if (time.youbi == '土' || time.youbi == '日') {
					//土日
					color = COLOR.HOLIDAY;
				}

				targetColorRange.setBackground(color);
			}
		}
		inputColor();
	}

	/**
	 * 今日の日付の色をつける
	 */
	function setToday() {
		var period = getPeriod();
		if (!period) return;

		//一旦全範囲をリセット
		sheet.getRange(YEAR_Y, AREA_ROOT_X, 3, period.width).setBackground(COLOR.NONE);

		var today = timeController.get();
		if (today.millisecond < period.start.millisecond || period.end.millisecond < today.millisecond) return;

		var diff = today.diffDay(period.start);
		sheet.getRange(YEAR_Y, AREA_ROOT_X + diff, 3, 1).setBackground(COLOR.TODAY);
	}


	return {
		isSpecialColumn: isSpecialColumn,

		isUpdate: isUpdate,
		updateCache: updateCache,
		update: update,

		getPeriod: getPeriod,

		getWidth: getWidth,
		getHeight: getHeight,

		setToday: setToday,
	};
})();


/**
 * タスク操作モジュール
 */
var row = (function() {

	var CACHE_NAME = 'row_cache';


	/**
	 * 開始日、作業日数、進捗の値が正しいかどうかチェック
	 * @param  {Date} start 
	 * @param  {number} time 
	 * @param  {number} progress 
	 * @return {boolean} 
	 */
	function validateInfo(start, time, progress) {
		if (!isDate(start) || !isNumber(time) || time <= 0) {
			return false;
		}
		return true;
	}

	/**
	 * 特定の列の開始日、作業日数、進捗の値をセルから取得
	 * @param  {number} y 
	 * @return {Object} {start: TimeInstance, time: number, progress: number}
	 */
	function getRowInfoFromCell(y) {
		var info = sheet.getRange(y, START_X, 1, 3).getValues()[0];
		var start = info[0];
		var time = Math.floor(info[1]);
		var progress = info[2] ? info[2] : 0;
		if (!validateInfo(start, time, progress)) {
			return null;
		}

		return {
			start: timeController.get(start),
			time: time,
			progress: progress,
		};
	}

	/**
	 * 特定の列の開始日、作業日数、進捗の値をキャッシュから取得
	 * @param  {number} y 
	 * @return {Object} {start: TimeInstance, time: number, progress: number}
	 */
	function getRowInfo(y) {
		if (!cache.isRow(y, CACHE_NAME)) return null;

		var temp = cache.getRow(y, CACHE_NAME);
		if (!temp) return null;

		return {
			start: timeController.get(temp.start),
			time: temp.time,
			progress: temp.progress,
		};
	}

	/**
	 * 特定の列の開始日などのキャッシュを更新する
	 * @param  {number} y 
	 */
	function updateCache(y) {
		var info = getRowInfoFromCell(y);
		if (!info) {
			cache.setRow(y, CACHE_NAME, null);
			return;
		}

		var temp = {
			start: info.start.millisecond,
			time: info.time,
			progress: info.progress
		};
		cache.setRow(y, CACHE_NAME, temp);
	}

	/**
	 * 全ての列のキャッシュを更新する
	 */
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

	/**
	 * 列を更新する必要があるかどうか
	 * @param  {number} x 
	 * @param  {number} y 
	 * @return {boolean} 
	 */
	function isUpdate(x, y) {
		if (y < AREA_ROOT_Y) return false;
		if (x != START_X && x != PERIOD_X && x != PROGRESS_X) return false;

		var cacheInfo = getRowInfo(y);
		var info = getRowInfoFromCell(y);
		if (!cacheInfo && !info) {
			return false;
		}

		return true;
	}

	/**
	 * 列を更新する
	 * @param  {number} y 値がなければ全ての列を更新
	 */
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

		var backgoundColors = sheet.getRange(AREA_ROOT_Y, AREA_ROOT_X, 1, project.getWidth()).getBackgrounds()[0];
		infos.forEach(function(info, index) {
			if (!info) return;

			var start = info.start;
			var time = info.time;
			var progress = info.progress;
			var posY = y ? y : AREA_ROOT_Y + index;

			//プロジェクト期間より後ろの日付は無視
			if (projectPeriod.end.millisecond < start.millisecond) return;

			//プロジェクトの期間より前の日付
			if (start.millisecond < projectPeriod.start.millisecond) {
				toast('期間エラー', '項目の開始日がプロジェクトの開始日よりも前になっています。' + posY + '行目');
				return;
			}
			
			/**
			 * 色付け
			 */
			
			var count = 0;
			var finishedCount = 0;
			var finishedPos = Math.floor(time * (progress / 100));

			function isHolidayColor(value) {
				return value == COLOR.HOLIDAY || value == COLOR.NATIONAL_HOLIDAY;
			}

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
				++finishedCount;
				if (finishedCount > finishedPos) {
					sheet.getRange(posY, AREA_ROOT_X + diff).setBackground(COLOR.UNFINISHED);
				}
				else {
					sheet.getRange(posY, AREA_ROOT_X + diff).setBackground(COLOR.FINISHED);
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




/**
 * ソート用モジュール
 */
var sort = (function() {
	function getRange() {
		return sheet.getRange(AREA_ROOT_Y, 1, project.getHeight(), AREA_ROOT_X - 1 + project.getWidth());
	}

	/**
	 * タスクの開始日でソート
	 */
	function start() {
		getRange().sort([
			{column: 3},
			{column: 2},
		]);
	}

	/**
	 * 担当者と開始日でソート
	 */
	function staff() {
		getRange().sort([
			{column: 2},
			{column: 3},
		]);
	}

	return {
		start: start,
		staff: staff,
	};
})();




/**
 * ロックできたらcallbackを呼ぶ。
 * ロック取得失敗時にはエラー出力
 * @param  {Function} callback 
 */
function tryLock(callback) {
	if (lock.start()) {
		callback();
	}
	else {
		toast('処理エラー', '処理が追いついていません。少し待ってから操作して下さいね。');
	}
}

/**
 * プロジェクトの期間の祝日のキャッシュを更新する
 */
function updateHolidayCache() {
	var period = project.getPeriod();
	if (!period) return;

	holiday.update(period.start, period.end);
}

