var Statistics;
AddEvent(window, "global", function() { Statistics = new StatisticsClass(); } );

/*------------------------------------------------------------------------
	dynamically update statistics values

	Date:		Name:	Description:
	2/18/05		JEA		Creation
------------------------------------------------------------------------*/
function StatisticsClass() {
	var me = this;
	var _service = "../service.aspx?";
	var _visitsNode = DOM.GetNode("lblVisitsToday");
	var _dataSizeNode = DOM.GetNode("lblDataSize");
	var _dataNameNode = DOM.GetNode("lnkDataFile");
	var _dataSavedNode = DOM.GetNode("lblDataSaved");
	var _usersNode = DOM.GetNode("lblUserCount");
	var _timer = setInterval("Statistics.Refresh()", 15000);
	
	this.Refresh = function() {
		var request = new ServerCall("GetStatistics")
		request.Callback = RefreshResult;
		request.Service = _service;
		//Global.ProgressBar.Start("Updating statistics");
		request.Start();
	}
	function RefreshResult(statistics) {
		//Global.ProgressBar.Stop();
		if (statistics.Errors.length == 0) {
			_visitsNode.innerHTML = statistics.Visits;
			_usersNode.innerHTML = statistics.Users;
			_dataSizeNode.innerHTML = statistics.DataFile.Size;
			_dataNameNode.innerHTML = statistics.DataFile.Name;
			_dataSavedNode.innerHTML = statistics.DataFile.Saved;
		}
	}
}