/**
 * キャッシュモジュール
 */
var cache = (function() {
	var ITEMS = spreadsheet.getId() + sheet.getSheetId() + '_' + 'cache_items';
	var mPublicCache = CacheService.getPublicCache();

	var CACHE_TIME = 60 * 24 * 5; //maximum is 6 hours.
	var NORMAL_HEADER = 'cache_header_normal_';
	var ROW_HEADER = 'cache_header_row_';
	var COLUMN_HEADER = 'cache_header_column_';
	var CELL_HEADER = 'cache_header_cell_';
	var ITEM_FOR_PROLONG = 'cache_item_for_prolong';

	var mItems = {};
	var mIsEmptyOrExpired = true;
	var mBatchStartCount = 0;


	function getKey(key) {
		return NORMAL_HEADER + key;
	}
	function getRowKey(y, key) {
		return ROW_HEADER + y + '_' + key;
	}
	function getColumnKey(x, key) {
		return COLUMN_HEADER + x + '_' + key;
	}
	function getCellKey(x, y, key) {
		return CELL_HEADER + x + '_' + y + '_' + key;
	}

	function updateCache() {
		if (isBatch()) return;

		mItems[ITEM_FOR_PROLONG] = new Date().getTime();
		var items = Utilities.jsonStringify(mItems);
		mPublicCache.put(ITEMS, items, CACHE_TIME);
	}



	function isEmptyOrExpired() {
		return mIsEmptyOrExpired;
	}

	function is(key) {
		return !isUndefined(mItems[key]);
	}

	function get(key) {
		if(is(key)) {
			return mItems[key];
		}
		else {
			null;
		}
	}
	function set(key, value) {
		mItems[key] = value;
		mIsEmptyOrExpired = false;
		updateCache();
	}
	function remove(key) {
		delete mItems[key];
		var items = Utilities.jsonStringify(mItems);
		mPublicCache.put(ITEMS, items, CACHE_TIME);
	}

	function removeAll() {
		var items = Utilities.jsonStringify({});
		mPublicCache.put(ITEMS, items, CACHE_TIME);
	}

	function prolong() {
		updateCache();
	}

	function checkCell(x, y) {
		return isNumber(x) && isNumber(y);
	}

	function startBatch() {
		++mBatchStartCount;
	}
	function finishBatch() {
		if (mBatchStartCount == 0) return;

		--mBatchStartCount;
		if (mBatchStartCount == 0) {
			updateCache();
		}
	}
	function isBatch() {
		return mBatchStartCount > 0;
	}

	function initialize() {
		var items = mPublicCache.get(ITEMS);
		if (!items) return;
		
		mIsEmptyOrExpired = false;
		mItems = Utilities.jsonParse(items);
	}
	initialize();


	return {
		is: function(key) {
			if (!key) return false;
			return is(getKey(key));
		},
		get: function(key) {
			if (!key) return null;
			return get(getKey(key));
		},
		set: function(key, value) {
			if (!key) return;
			set(getKey(key), value);
		},
		remove: function(key) {
			if (!key) return;
			remove(getKey(key));
		},

		isRow: function(y, key) {
			if (!key || !isNumber(y)) return false;
			return is(getRowKey(y, key));
		},
		getRow: function(y ,key) {
			if (!key || !isNumber(y)) return null;
			return get(getRowKey(y, key));
		},
		setRow: function(y, key, value) {
			if (!key || !isNumber(y)) return;
			set(getRowKey(y, key), value);
		},
		removeRow: function(y, key) {
			if (!key || !isNumber(y)) return;
			remove(getRowKey(y, key));
		},


		isColumn: function(key) {
			if (!key || !isNumber(x)) return false;
			return is(getColumnKey(x, key));
		},
		getColumn: function(x ,key) {
			if (!key || !isNumber(x)) return null;
			return get(getColumnKey(x, key));
		},
		setColumn: function(x, key, value) {
			if (!key || !isNumber(x)) return;
			set(getColumnKey(x, key), value);
		},
		removeColumn: function(x, key) {
			if (!key || !isNumber(x)) return;
			remove(getColumnKey(x, key));
		},


		isCell: function(x, y, key) {
			if (!key || !isCell(x, y)) return false;
			return is(getCellKey(key));
		},

		getCell: function(x, y, key) {
			if (!key || !isCell(x, y)) return null;
			return get(getCellKey(x, y, key));
		},

		setCell: function(x, y, key, value) {
			if (!key || !isCell(x, y)) return;
			set(getCellKey(x, y, key), value);
		},
		removeColumn: function(x, key) {
			if (!key || !isCell(x, y)) return;
			remove(getCellKey(x, y, key));
		},

		removeAll: removeAll,

		prolong: prolong,

		isNullOrExpired: isEmptyOrExpired,

		startBatch: startBatch,
		finishBatch: finishBatch,
		isBatch: isBatch,
	};
})();


