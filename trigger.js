

/**
 * トリガー操作モジュール
 */

var trigger = (function() {

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
