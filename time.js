
/**
 * 日付を扱うクラスを生成する
 * @return {TimeInstance} TimeInstance
 */
var timeController = (function() {
	return  {
		instance: function(obj_1, obj_2, obj_3) {
			if (isNull(obj_1)) {
				return null;
			}

			if(obj_1 === '') {
				return null;
			}

			var dateOrTimeInstance = obj_1;
			if (!isUndefined(obj_3)) {
				//year, month and day
				dateOrTimeInstance = new Date(obj_1, obj_2 - 1, obj_3);
			}
			else if (obj_2 !== undefined) {
				//x, y
				var date = sheet.getRange(obj_2, obj_1).getValue();
				if (!isDate(date)) {
					return null;
				}
				dateOrTimeInstance = date;
			}
			return new _TimeInstance(dateOrTimeInstance);
		},

		get: function(obj_1, obj_2, obj_3) {
			return timeController.instance(obj_1, obj_2, obj_3).get();
		},
	};
})();

/**
 * 日付情報を持ったオブジェクト
 */
var TimeInstance = function(millisecond) {
	var d = new Date(millisecond);

	this.year = d.getFullYear();
	this.month = d.getMonth() + 1;
	this.day = d.getDate();
	this.hour = d.getHours();
	this.minute = d.getMinutes();
	this.second = d.getSeconds();
	this.millisecond = d.getTime();
	this.youbi = YOUBI[d.getDay()];
	this.date = d;
};
(function(p) {
	var ADJUST = 1000;

	p.diffDay = function(targetTimeInstance) {
		if (this.millisecond > targetTimeInstance.millisecond) {
			return Math.floor((this.millisecond + ADJUST - targetTimeInstance.millisecond) / DAY);
		}
		else {
			return Math.floor((targetTimeInstance.millisecond + ADJUST - this.millisecond) / DAY);
		}
	};
})(TimeInstance.prototype);


/**
 * 日付を操作できるクラス
 * get()でフわかりやすい形で年月日などを取得できる。
 * @param {any} obj empty or Date or _TimeInstance or TimeInstance or millisecond.
 */
var _TimeInstance = function(obj) {
	this.millisecond = 0;
	if (isDate(obj)) {
		this.millisecond = obj.getTime();
	}
	else if (obj instanceof _TimeInstance || obj instanceof TimeInstance) {
		this.millisecond = obj.millisecond;
	}
	else if (obj) {
		this.millisecond = obj;
	}
	else {
		this.millisecond = new Date().getTime();
	}
};
(function(p) {
	/**
	 * 日にちをプラスする
	 * @param  {number} num 日数
	 */
	p.plusDay = function(num) {
		this.millisecond += num * DAY;
		return this;
	};

	/**
	 * 日にちをマイナスする
	 * @param  {number} num 日数
	 */
	p.minusDay = function(num) {
		this.plusDay(-num);
		return this;
	};

	p.get = function() {
		return new TimeInstance(this.millisecond);
	};

})(_TimeInstance.prototype);







