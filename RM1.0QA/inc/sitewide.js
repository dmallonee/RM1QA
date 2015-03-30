var currentSection;

//onMouseOver="mouseOver(this);" onMouseOut="mouseOut('all',this);"
function mouseOver(theObj, theSectionName) {
	if (theSectionName != currentSection) {
		theObj.src = '/images/tab_' + theSectionName + '_sel.gif';
	}
}

function mouseOut(theObj, theSectionName) {
	if (theSectionName != currentSection) {
		theObj.src = '/images/tab_' + theSectionName + '.gif';
	}
}

function highlightSection(theSectionName) {
    currentSection = theSectionName;
    
	if (document.getElementById) {
		var targetElement = document.getElementById('menuIcon_' + theSectionName);
		targetElement.src = '/images/tab_' + theSectionName + '_sel.gif';
	}
}

function cl(inp, val) {
	if (inp.value == val) inp.value = "";
}

function fl(inp, val) {
	if (inp.value == "") inp.value = val;
}
