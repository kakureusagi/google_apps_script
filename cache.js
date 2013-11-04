/**
 * キャッシュモジュール
 */
var cache = (function() {
	var ITEMS = spreadsheet.getId() + sheet.getSheetId() + '_' + 'cache_items';
	var mPublicCache = CacheService.getPublicCache();

	var CACHE_TIME = 60 * 24 * 6; //maximum is 6 hours.
	var NORMAL_HEADER = 'cache_header_normal_';
	var ROW_HEADER = 'cache_header_row_';
	var COLUMN_HEADER = 'cache_header_column_';
	var CELL_HEADER = 'cache_header_cell_';

	var mItems = {};


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
		var items = Utilities.jsonStringify(mItems);
		mPublicCache.put(ITEMS, items, CACHE_TIME);
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
		var items = Utilities.jsonStringify(mItems);
		mPublicCache.put(ITEMS, items, CACHE_TIME);
	}

	function checkCell(x, y) {
		return isNumber(x) && isNumber(y);
	}

	function initialize() {
		var items = mPublicCache.get(ITEMS);
		if (!items) return;
		
		mItems = Utilities.jsonParse(items);
		Logger.log('cache items is %s.', mItems);
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
	};
})();


