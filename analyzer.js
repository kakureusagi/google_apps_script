/**
 * 時間計測クラス
 */
var Analyzer = function(topMessage) {
	this.times = [timeController.get()];
	this.topMessage = topMessage ? topMessage : '';
	this.count = 0;
};
(function(p) {

	p.set = function (message) {
		message = message ? message : '';
		var current = timeController.get();
		var before = this.times[this.times.length - 1]
		Logger.log((current.millisecond - before.millisecond) + 'ms: ' + this.topMessage + '[' + this.count + '] => ' + message);
		this.times.push(current);
		++this.count;

		return this;
	};

	p.back = function() {
		if (this.count > 0) {
			--this.count;
			--this.times.length;
		}
		return this;
	};

})(Analyzer.prototype);

