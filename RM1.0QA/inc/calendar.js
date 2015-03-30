function Calendar() {
  //var formNumber = request.getParameter("formNumber");
  //var formName = request.getParameter("formName");
  //var formRef = (formName) ? window.opener.document.getElementById(formName) : null;
  var allowPast =  false //(formRef) ? eval(formRef.dateboxPast.value) : null;
  var maxMonths = 13 // (formRef) ? parseInt(formRef.maxMonths.value) : null;
  var secondDate = true // (formRef) ? eval(formRef.secondDate.value) : null;
  var dot = '<span class="bold">&middot;</span>';

  // Server-side hook
  var currentDate = cd = (window.serverTime) ? new Date(serverTime[0],serverTime[1],serverTime[2]) : new Date();
  var displayDate = dd = (window.serverTime) ? new Date(serverTime[0],serverTime[1],serverTime[2]) : new Date();
  var maxYears = Math.ceil(maxMonths / 12) + ((allowPast)?1:0);
  var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var daysPerMonth = [31,28,31,30,31,30,31,31,30,31,30,31];
  var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  
  this.leapYear = function(yr) {
    if ((yr/4) != Math.floor(yr/4)) return false;
    if ((yr/100) != Math.floor(yr/100)) return true;
    if ((yr/400) != Math.floor(yr/400)) return false;
    return true;  
  }
  
  this.testLeap = function() {
    daysPerMonth[1] = (this.leapYear(dd.getFullYear())) ? 29 : 28;
  }
  
  this.writeYear = function() {
    var yr = cd.getFullYear() - ((allowPast)?1:0);
    var yearCont = "";
    for (i=0; i < maxYears; i++) {
      var isDispYr = (yr == dd.getFullYear()) ? true : false;
      yearCont += dot + ((!isDispYr) ? '<a href="javascript:cal.setCalDate('+yr+',0,1)">' : '') + yr + ((!isDispYr) ? '</a>' : '') + "&nbsp;";
      yr++;
    }
    csspObj('Year').setProperty('innerHTML', yearCont);
  }
  
  this.writeMonth = function() {
    var remMnths = (maxMonths + cd.getMonth()) - ((dd.getFullYear() - cd.getFullYear())*12);
    var dispMths = Math.min(remMnths, months.length);
    var mthCont = '<table width="266" border="0" cellspacing="0" cellpadding="2"><tr>';

    for (var i=0; i < 12; i++) {
      var isThisMth = (i == dd.getMonth()) ? true : false;
      var showLink = (!isThisMth && (((i >= cd.getMonth()) && (dd.getFullYear() == cd.getFullYear())) || (dd.getFullYear() > cd.getFullYear()))) ? true : false;
      if (allowPast && !isThisMth) showLink = true;
      var s = (showLink) ? '<a href="javascript:cal.setCalDate('+dd.getFullYear()+','+i+',1)">' : '';
      var e = (showLink) ? '</a>' : '';
      
      if (i == 6) mthCont += '</tr><tr>';
      var mthDisp = (i <= (dispMths-1)) ? dot+s+months[i].substring(0,3)+e : "&nbsp;";
      mthCont += '<td width="40">'+mthDisp+'</td>';
    }  
    mthCont += '</tr></table>';
    csspObj('Month').setProperty('innerHTML', mthCont);
  }
  
  this.writeDate = function() {
    // First day of the display month
    var tempDate = new Date(dd.getFullYear(), dd.getMonth(), 1);
    var startDay = tempDate.getDay();
    var moDay = months[dd.getMonth()] + " " + dd.getFullYear();
    var dateCont = '<table width="266" border="0" cellspacing="0" cellpadding="2" bgcolor="#CED7E7"><tr><td class="bold">&nbsp;&nbsp;'+moDay+'</td></tr></table>';
    dateCont += '<table width="266" height="125" border="0" cellspacing="0" cellpadding="1"><tr>';
    
    // Weekday headers
    for (var i=0; i<days.length; i++) {
      var isHiDay = (i == 0 || i == 6) ? "hilite" : "bold";
      dateCont += '<td class="'+isHiDay+'" width="38" align="center">'+days[i].substring(0,1)+'</td>';
    }
  
    dateCont += '</tr><tr>';
    var dispDate = 1;
    // Dates
    for (var j=0; j<42; j++) {
      if ((j%7 == 0) && j != 0) dateCont += '</tr><tr>';
      if (j >= startDay && dispDate <= daysPerMonth[dd.getMonth()]) {
        var hilite = (dispDate == dd.getDate()) ?  true : false;
        var showLink = ((dispDate >= dd.getDate()) && (dd.getFullYear() >= cd.getFullYear())) ? true : false;
        if (allowPast) showLink = true;
        var hs = (hilite) ? '<table cellspacing="0" cellpadding="0" border="1" width="31" height="20" class="outline" bordercolor="#cccccc"><tr><td class="bold" align="center">': '';
        var he = (hilite) ? '</td></tr></table>' : '';
        var ls = (showLink) ? '<a href="javascript:cal.sendDate('+dd.getFullYear()+','+dd.getMonth()+','+dispDate+')">' : '';
        var le = (showLink) ? '</a>' : '';
        dateCont += '<td width="38" align="center">'+hs+ls+dispDate+le+he+'</td>' 
        dispDate++;
      } else {
        dateCont += '<td width="38">&nbsp;</td>';
      }
    }
    dateCont += '</tr></table>';
    csspObj('Date').setProperty('innerHTML', dateCont);
  }
  
  this.setCalDate = function(y, m, d) {
    if ((y == cd.getFullYear()) && (m < cd.getMonth()) && !allowPast) m = cd.getMonth();
    if ((y == cd.getFullYear()) && (cd.getMonth() == m)) d = cd.getDate();
    var tempD = new Date(y, m, d);
    dd = tempD;

    this.writeCal();
  }
  
  this.writeCal = function() {
    with (this) {
      testLeap();
      writeDate();
      writeMonth();
      writeYear();
    }
  }
  
  this.sendDate = function(y, m, d) {
    if (formRef) {
      var yr = y.toString().substring(2);
      var mo = m+1;
      var da = d;
      if (mo < 9) mo = "0" + mo;
      if (da < 9) da = "0" + da;
      
      formRef["iyear" + formNumber].value = yr; 
      formRef["imonth" + formNumber].value = mo; 
      formRef["iday" + formNumber].value = da; 
      
      if (!(formNumber % 2)) {
        var secondForm = parseInt(formNumber) + 1;
        var secDt = new Date(y, m, (d+1));
        
        yr = secDt.getFullYear().toString().substring(2);
        mo = secDt.getMonth()+1;
        da = secDt.getDate();
        
        if (mo < 9) mo = "0" + mo;
        if (da < 9) da = "0" + da;
        
        formRef["iyear" + secondForm].value = yr; 
        formRef["imonth" + secondForm].value = mo; 
        formRef["iday" + secondForm].value = da; 
      }
      window.close();
    }
  }
}

var request = new Object();
request.getParameter = function(param) {
  var s = window.location.search;
  if(!s) return null;
  if(!(s.indexOf(param+'=')+1)) return null;
  return s.split(param+'=')[1].split('&')[0];
}

function calInit() {
  cal = new Calendar();
  cal.writeCal();
}

window.onloadHandlers[onloadHandlers.length] = "calInit()";