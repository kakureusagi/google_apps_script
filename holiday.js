

/**
 * 祝日管理モジュール
 */
var holiday = (function() {

	var CACHE_NAME = '_holiday_times_cache_';

	function getTimes() {
		var holidayTimes = [];
		if (!cache.is(CACHE_NAME)) return holidayTimes;

		var holidayMilliseconds = cache.get(CACHE_NAME);
		for (var i = 0 ; i < holidayMilliseconds.length ; ++i) {
			holidayTimes.push(timeController.get(holidayMilliseconds[i]));
		}

		return holidayTimes;
	}

	function update(start, end) {
		var holidays = CalendarApp
			.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com")
			.getEvents(
				start.date,
				end.date
			);

		var holidayTimes = [];
		for (var i = 0 ; i < holidays.length ; ++i) {
			holidayTimes.push(timeController.get(holidays[i].getStartTime()).millisecond);
		}

		cache.set(CACHE_NAME, holidayTimes);
	}

	return {
		getTimes: getTimes,
		update: update,
	};
})();


