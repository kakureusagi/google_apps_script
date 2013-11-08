

/**
 * ロックモジュール
 */
var lock = (function() {

	var publicLock = LockService.getPublicLock();
	var MAX_TIME = 30000;

	function start(millisecond) {
		return publicLock.tryLock(millisecond ? millisecond : MAX_TIME);
	}

	return {
		start: start,
	};
})();


