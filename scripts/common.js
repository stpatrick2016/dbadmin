window.onerror = function(){;};
var imageCache = new Image();
imageCache.src = "images/icon_show.gif";

//language defined strings
//var langShow = "Show";
//var langHide = "Hide"

function newTableToggle(idTD, idImg){
	var td = document.getElementById(idTD);
	var img = document.getElementById(idImg);
	if(td != null && img != null){
		var isHidden = td.style.display == "none" ? true : false;
		img.src = isHidden ? "images/icon_hide.gif" : "images/icon_show.gif";
		img.alt = isHidden ? langHide : langShow;
		td.style.display = isHidden ? "" : "none";
	}
}

function DBA_popupWindow(url, target, width, height){
	var features;
	features = 'location=0,menubar=0,scrollbars,resizable,dependent,status=0,toolbar=0,width=' + width + ',innerWidth=' + width + ',height=' + height + ',innerHeight=' + height;
	if (window.screen) {
		var ah = screen.availHeight - 30;
		var aw = screen.availWidth - 10;

		var xc = (aw - width) / 2;
		var yc = (ah - height) / 2;

		features += ",left=" + xc + ",screenX=" + xc;
		features += ",top=" + yc + ",screenY=" + yc;
	}
	var wnd = window.open(url, target, features);
	wnd.focus();
}
function copyToClipboard(id){
	// Copy displayed code to user's clipboard.
	var textRange = document.body.createTextRange();
	textRange.moveToElementText(id);
	textRange.execCommand("Copy");
	alert("All necessary code has been copied to your clipboard");
}
