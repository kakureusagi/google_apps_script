/**
 * キャッシュクラス
 */
var cache = (function() {
	var ITEMS = 'cache_items_';
	var NORMAL_HEADER = 'cache_header_normal_';
	var CELL_HEADER = 'cache_header_x_y_';

	var mItems = {};
	var mPublicCache = CacheService.getPublicCache();



	function getKey(key) {
		return NORMAL_HEADER + key;
	}
	function getCellKey(x, y, key) {
		return CELL_HEADER + x + '_' + y + '_' + key;
	}


	function get(key) {
		if(mItems[key]) {
			return mItems[key];
		}
		else {
			null;
		}
	}
	function set(key, value) {
		mItems[key] = value;
		var items = Utilities.jsonStringify(mItems);
		mPublicCache.put(ITEMS, items);
	}

	function isNumber(x, y) {
		return !isNaN(x) && !isNaN(y);
	}

	function initialize() {
		var items = mPublicCache.get(ITEMS);
		if (!items) return;
		
		mItems = Utilities.jsonParse(items);
	}
	initialize();


	return {
		get: function(key) {
			if (!key) return null;
			return get(getKey(key));
		},

		set: function(key, value) {
			if (!key) return;
			set(getKey(key), value);
		},

		getCell: function(x, y, key) {
			if (isNumber(x, y)) return null;
			return get(getCellKey(x, y, key));
		},

		setCell: function(x, y, key, value) {
			if (isNumber(x, y)) return;
			set(getCellKey(x, y, key), value);
		}
	};
})();


