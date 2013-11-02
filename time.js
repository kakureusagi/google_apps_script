

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
	p.diffDay = function(target) {
		if (this.millisecond > target.millisecond) {
			return Math.floor((this.millisecond + HOUR - target.millisecond) / DAY);
		}
		else {
			return Math.floor((target.millisecond + HOUR - this.millisecond) / DAY);
		}
	};
})(TimeInstance.prototype);



/**
 * 日付を操作できるクラス
 * get()でフわかりやすい形で年月日などを取得できる。
 * @param {any} obj empty or Date or TimeInstance or TimeInstance.get().
 */
var _TimeInstance = function(obj) {
	this.millisecond = 0;
	if (obj instanceof Date) {
		this.millisecond = obj.getTime();
	}
	else if (obj instanceof _TimeInstance || obj instanceof TimeInstance) {
		this.millisecond = obj.millisecond;
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



/**
 * 日付を扱うクラスを生成する
 * @return {TimeInstance} TimeInstance
 */
var timeController = (function() {
	return  {
		instance: function(obj_1, obj_2, obj_3) {
			if (obj_1 === null) {
				return null;
			}

			if(obj_1 === '') {
				return null;
			}

			var dateOrTimeInstance = obj_1;
			if (obj_3 !== undefined) {
				dateOrTimeInstance = new Date(obj_1, obj_2 - 1, obj_3);
			}
			else if (obj_2 !== undefined) {
				var date = sheet.getRange(obj_2, obj_1).getValue();
				if (!date instanceof Date) {
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




