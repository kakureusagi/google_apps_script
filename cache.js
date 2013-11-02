/**
 * キャッシュクラス
 */
var cache = (function() {
	var HEADER = 'cache_header_';
	var CELL_HEADER = 'cache_header_x_y_';

	var publicCache = CacheService.getPublicCache();



	function getKey(key) {
		return HEADER + key;
	}
	function getCellKey(x, y, key) {
		return CELL_HEADER + x + '_' + y + '_' + key;
	}


	function get(key) {
		var temp = publicCache.get(key);
		if (!temp) return null;

		return Utilities.jsonParse(temp);
	}
	function set(key, value) {
		var temp = Utilities.jsonStringify(value);
		publicCache.put(key, temp);
	}

	function isNumber(x, y) {
		return !isNaN(x) && !isNaN(y);
	}



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


