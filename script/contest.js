var Contest;
AddEvent(null, "global", function() { Contest = new ContestClass(); } );

/*------------------------------------------------------------------------
	process contest activities

	Date:		Name:	Description:
	2/18/05		JEA		Creation
------------------------------------------------------------------------*/
function ContestClass() {
	var me = this;
	var _edit = (DOM.GetNode("fldStartDate") != null);
	var _mediaType = 2115;		//bitmask of media File.Types
	var _projectType = 1128;	//bitmask of project File.Types
	
	if (_edit) {
		var _fileList = DOM.GetNode("ampFileList");
		var _startDate = DOM.GetNode("fldStartDate");
		var _endVoteDate = DOM.GetNode("fldEndVoteDate");
		var _stopDate = DOM.GetNode("fldStopDate");
		var _plugins = DOM.GetNode("fldFreePlugins");
		var _media = DOM.GetNode("fldGeneratedMediaOnly");
		var _project = DOM.GetNode("forProjects", true);
		var _winners = DOM.GetNode("fldWinners");
		var _prizes = DOM.GetNode("tbPrizes");
		SetupEditValidation();
		SetFileOptions();
	} else {
		var _ballot = DOM.GetNode("ampBallot");
		var _voteNodes = DOM.GetElementsByID(/ampBallot_/, _ballot);
		SetupBallotValidation();
	}
	// auto-select all file types
	this.SelectAllFiles = function() {
		for (var x = 0; x < _fileList.options.length; x++) {
			_fileList.options[x].selected = true;
		}
		SetFileOptions();
	}
	// contest selection
	this.Change = function(node) {
		var id = node.options[node.selectedIndex].value;
		location.href = "contests.aspx" + ((id.length == 36) ? "?id=" + id : "");
	}
	// respond to file selection
	this.FileClick = function(node) { SetFileOptions(); }
	
	// prevent duplicate votes
	this.CheckBallot = function(node) {
		var thatVote, vote;
		var thisVote = node.options[node.selectedIndex].value;
		var re = /_(\d+)$/;
		var rank = (re.test(node.id)) ? re.exec(node.id)[1] : 0;
		
		for (var x = 0; x < _voteNodes.length; x++) {
			vote = _voteNodes[x];
			thatVote = vote.options[vote.selectedIndex].value;
			if ((rank != x + 1) && thatVote == thisVote) {
				vote.selectedIndex = 0;
			}
		}
	}
	// check file type to enable/disable checkboxes
	function SetFileOptions() {
		var enabled = ((Selected() & _projectType) > 0);
		_project.style.display = enabled ? "block" : "none";
		//DOM.SetOpacity(_project, (enabled ? 100 : 0));
		if (!enabled) { _plugins.checked = _media.checked = false; }
	}
	// 
	function Selected() {
		var selected = 0;		//bitmask
		for (var x = 0; x < _fileList.options.length; x++) {
			if (_fileList.options[x].selected) { selected |= _fileList.options[x].value; }
		}
		return selected;
	}
	
	// validate edit form
	function SetupEditValidation() {
		if (typeof(Validation) == "undefined") { setTimeout(SetupEditValidation, 100); return; }
		
		// validate date orders
		Validation.Functions.push( function() {
			var start = Validation.GetDate(_startDate);
			var voteEnd = Validation.GetDate(_endVoteDate);
			var end = Validation.GetDate(_stopDate);
			if (start && voteEnd && end) {
				if (start >= end) {	return "End date must be after start date";	}
				if (start >= voteEnd) { return "Finish voting date must be after start date"; }
				if (voteEnd > end) { return "Finish voting date must be on or before end date"; }
			}
			return null;
		} );
		// prevent mixing of media and project file types
		Validation.Functions.push( function() {
			var selected = Selected();
			if ((selected & _projectType) > 0 && (selected & _mediaType) > 0) {
				return "Cannot allow both project and media file types";
			}
			return null;
		} );
		// ensure prize count equals winner count
		Validation.Functions.push( function() {
			if (_prizes.value.length > 1) {
				// first split is null string, so substract
				var prizes = _prizes.value.split("-");
				if ((prizes.length - 1) != _winners.value) {
					return "The number of prizes does not match the number of winners";
				}
			}
			return null;
		} );
		
	}
	
	// validate ballot
	function SetupBallotValidation() {
		if (typeof(Validation) == "undefined") { setTimeout(SetupBallotValidation, 100); return; }
		
		// validate that vote was cast
		Validation.Functions.push( function() {
			var vote;
			for (var x = 0; x < _voteNodes.length; x++) {
				vote = _voteNodes[x];
				if (vote.options[vote.selectedIndex].value.length == 36) {
					return null;
				}
				return "At least one entry on the ballot must be voted for";
			}
		} );
	}
}