function openWindow(url) { popupWin = window.open(url,'new_page','width=400,height=400') }
function openWindow2(url) { popupWin = window.open(url,'new_page','width=400,height=450') }
function openWindow3(url) { popupWin = window.open(url,'new_page','width=400,height=450,scrollbars=yes') }
function openWindow4(url) { popupWin = window.open(url,'new_page','width=400,height=525') }
function openWindow5(url) { popupWin = window.open(url,'new_page','width=450,height=525,scrollbars=yes,toolbars=yes,menubar=yes,resizable=yes') }
function openWindow6(url) { popupWin = window.open(url,'new_page','width=500,height=450,scrollbars=yes') }
function openWindowHelp(url) { popupWin = window.open(url,'new_page','width=470,height=200,scrollbars=yes') }

function unsub_confirm(link) {
var where_to = confirm("Do you really want to Unsubscribe?");
if (where_to == true) { popupWin = window.open(link,'new_page','width=400,height=400') } }

function jumpTo(s) {if (s.selectedIndex != 0) location.href = s.options[s.selectedIndex].value;return 1;}
function setDays() {document.DaysFilter.submit(); return 0;}

function ChangePage(fnum) {if (fnum == 1) {document.PageNum1.submit();} else {document.PageNum2.submit();} }

function autoReload() {document.ReloadFrm.submit()}
function SetLastDate() {document.LastDateFrm.submit()}
