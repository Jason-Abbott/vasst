AddEvent(window, "load", SelectStyle);

function SelectStyle() {
	for (var x = 0; x < document.forms.length; x++) {
		for (var y = 0; y < document.forms[x].elements.length; y++) {
			var field = document.forms[x].elements[y];
			if (field.type == "select-one" || field.type == "select-multiple") {
				for (var z = 0; z < field.options.length; z++) {
					//alert(field.options[z].text);
					if (field.options[x].selected) {
						field.options[x].style.color = "#ff0000";
					}
				}
				AddEvent(field, "click", Talk);
			}
		}
	}
}

function Talk() {
	alert("here");
}