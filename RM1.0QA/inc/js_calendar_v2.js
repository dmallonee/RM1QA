//based on code from visionmonster.com
var o_navigator = navigator.userAgent.toLowerCase();
var isMacIE = (o_navigator.indexOf("msie 5")>-1&&o_navigator.indexOf("mac")>-1) ? 1 : 0;
var isPCIE = (o_navigator.indexOf("msie")>-1&&!isMacIE&&o_navigator.indexOf("opera")==-1) ? 1 : 0;
var isOpera = o_navigator.indexOf("opera")>-1 ? 1 : 0;
var mArray	= new Array("January","February","March","April","May","June","July","August","September","October","November","December");
var dArray 	= new Array("Su","Mo","Tu","We","Th","Fr","Sa");
var datesArray 		= new Array(31,28,31,30,31,30,31,31,30,31,30,31);
var today 			= new Date();			//todays date
var cD 				= today.getDay();		//current day of week 0-6
var cT				= today.getDate();		//current day 1-31
var cM				= today.getMonth();		//current month 0-11
var cMDs 			= datesArray[cM];		//number of days in current month
var cY				= today.getFullYear();	//js method	not used here//current Year
var newT			= cT;					//cal selected date
var newM			= cM;					//cal selected month
var newY			= cY;					//cal selected year
var newDs			= cMDs;					//days for selected Month
var newD			= cD;					//day of week
var numCalendars	= 2;					//number of calendars you want to create
var calDirection	= "vertical";			//put "horizontal" or "vertical"
var calopen 		= 0;					//boolean set state of iframe "0" closed "1" open;
var o_windowparent  = top;					//set parent frame
var o_input 		= 0;
var o_month = 0; var o_date = 0;
var o_iframecal = 0;
var o_from;var o_to;var v_from;var v_to;
var d_makefrom=0; var d_maketo=0;var s_lang="us";
var o_currentDate = false;
var i_firstYear = cY; var i_firstMonth = cM;
var o_row = null;
var i_numcal = 0;
var o_calbody = null; //where to write the calendar
var s_closecal = "<div class='calClose'><a href='#' onclick='top.closeCal();return false;'>close</a>Select a Date:</div>";
var b_date331 = 0;
var a_input = 0;
var o_parent;
var s_inputtype = "object"; //for text input or select list input
var a_v_input = null;
function findFirstDay(){
	firstDay = new Date();
	firstDay.setDate(1);
	firstDay.setMonth(newM);
	firstDay.setFullYear(newY);
	return firstDay.getDay();
}
function check331(d_date){
	i_date331 = Math.floor((d_date-today)/86400000);
	b_check331 = (i_date331>330) ? 1 : 0;
	return b_check331;
}
function vm_setupCal(){
	i_numcal = 0;
	vm_makeCal(cM);
}
function preventClose(evt){
	if(isOpera)evt.stopPropagation()
}
function ty_makeDate(which){
	d_makedate = new Date(newY,newM,which);
	b_date331 = check331(d_makedate);
	if((cT>which && cM == newM && cY == newY)||b_date331){
		s_makeDate = "<td class='calDateOff'>";
		s_makeDate+= which;
	}else{
		if(d_makefrom||d_maketo){
			s_makeDate = ((d_makedate.toString()==d_makefrom.toString())||(d_makedate.toString()==d_maketo.toString())) ? "<td class='calDateSel'>" : (d_makedate>d_makefrom&&d_makedate<d_maketo&&d_makefrom) ? "<td class='calDateRng'>" : "<td class='calDate'>";
		}else{
			s_makeDate = "<td class='calDate'>";
		}
		s_makeDate+= "<a href='#' onclick='top.ty_setDate("+newM+","+which+",this.parentNode,"+newY+");return false;' class='calDateA'>";
		s_makeDate+= which;
		s_makeDate+="</a>";
	}
	s_makeDate+="</td>\n";
	return s_makeDate;
}
function ty_maketr(what){
	s_tr = "<tr>\n";
	s_tr+= what;
	s_tr+= "</tr>\n";
	return s_tr;
}
function ty_changeMonths(which){
	i_numcal = 0;
	o_calbody.innerHTML ="";
	if(which < 0){
		which=11;
		newY--;
	}
	vm_makeCal(which);
}
function vm_makeCal(whichMonth){
	o_cal= "";
	o_caltr="";o_caltd="";
	newM = whichMonth;
	if(newM < cM) newY = cY+1;
	if (newM>=12){
		newM=whichMonth-12;
		newY++;
	}
	if(i_numcal==0)i_firstMonth  = newM;
	newDs = datesArray[newM];
	isLeap 	= (newY % 4 == 0 && (newY % 100 !=0 || newY % 400 ==0 )) ? 1:0
	if (newM==1) newDs=newDs+isLeap;
	newD = findFirstDay();
	countDay = newD;
	s_calclass = (calDirection=="vertical")? "calTableV" : "calTableH";
	o_cal+="<table month='"+newM+"' year='"+newY+"' cellpadding='0' cellspacing='0' border='0' class='"+s_calclass+"'>\n";
	o_caltr+= "<tr class='calRowHighlight'>\n";
	o_caltd+= "<td colspan='7' class='calLabel'>";
	o_caltd+= mArray[newM]+"&nbsp;"+newY;
	o_caltd+= "</td>";
	o_caltr+=o_caltd;
	o_caltr+="</tr>\n";
	o_cal+=o_caltr;
	o_caltd = "";
	for(i=0;i < dArray.length;i++){
		o_caltd+="<td class='calDayName'>";
		o_caltd+=dArray[i];
		o_caltd+="</td>\n";
	}
	o_caltr = ty_maketr(o_caltd);
	o_cal+=o_caltr;
	o_caltd = "";
	i_calRows = 0;
	for (d=1;d<=newDs;d++){
		if(d==1)for(bd=0;bd < newD;bd++)o_caltd += "<td class='calDate'>&nbsp;</td>\n";
		o_caltd += ty_makeDate(d);
		countDay++;
		if(countDay==7){
			countDay=0;
			o_caltr = ty_maketr(o_caltd);
			o_cal+=o_caltr;
			o_caltd = "";
			i_calRows++
		}
		if(d==newDs && countDay!=0){
			for (bd=countDay;bd < 7;bd++) o_caltd += "<td class='calDate'>&nbsp;</td>\n";
			o_caltr = ty_maketr(o_caltd);
			o_cal+=o_caltr;			
			o_caltd ="";
			i_calRows++
		}
	}
	if(i_calRows < 6){
		o_caltd = "";
		for(bd=0;bd < 7;bd++) o_caltd += "<td class='calDate'>&nbsp;</td>\n";
		o_caltr = ty_maketr(o_caltd);
		o_cal+= o_caltr;
	}
	o_cal+="</table>";
	o_calbody.innerHTML += (i_numcal==0) ? (newM==cM&&newY==cY) ? s_closecal+"<span class='calNavA'>&nbsp;</span>" : s_closecal+"<a href='#' onclick='top.ty_changeMonths("+i_firstMonth+"-1);top.preventClose(event);return false' class='calNavA'>previous month</a>" : "";
	o_calbody.innerHTML += o_cal;
	i_numcal++;
	if(i_numcal==numCalendars&&!(newM==cM-1)&&!b_date331)o_calbody.innerHTML+= "<a href='#' onclick='top.ty_changeMonths("+i_firstMonth+"+1);top.preventClose(event);return false;' class='calNavA'>next month</a>"
	if(i_numcal < numCalendars)vm_makeCal(newM+1);
	else if (i_firstMonth > newM){
		newY--;
	}
}
function ty_setDate(whatMonth,whatDate,whatTD,whatYear){
	o_currentDate = whatTD;
	o_currentDate.className = "calDateSel";
	if(typeof(o_input)=="object"){
		o_input.value = (s_lang=="us") ? (whatMonth+1)+"/"+whatDate+"/"+whatYear : whatDate+"/"+(whatMonth+1)+"/"+whatYear;
	}else{
		top.document.getElementById(a_v_input[0]).selectedIndex = whatMonth;
		top.document.getElementById(a_v_input[1]).selectedIndex = whatDate-1;
	}
	closeCal();
}
function hideCalendar(){
	o_caldiv.style.display = "none"
	if(o_parent) o_parent.className = "cbrow"
}
function splitDate(s_input, s_mode){
	this.delimitor = (s_input.indexOf("/")>-1) ? "/" : (s_input.indexOf(".")>-1) ? "." : (s_input.indexOf("-")>-1) ? "-" : (s_input.indexOf(",")>-1) ? "," : "/";
	a_input = s_input.split(this.delimitor);
	this.date = -1;this.month = -1;this.year = -1;
	if(a_input.length==3&&!isNaN(a_input[0])&&!isNaN(a_input[1])&&!isNaN(a_input[2])){
		this.month = (s_mode=="us") ? parseInt(a_input[0],10)-1 : parseInt(a_input[1],10)-1;
		this.date = (s_mode=="us") ? parseInt(a_input[1],10) : parseInt(a_input[0],10);
		this.year = a_input[2];
		if(this.month>11||this.month<0)this.month=-1;
		if(this.date>31||this.month<0)this.date=-1;
		i_yrlength = this.year.toString().length;
		if(i_yrlength==2)this.year = "20"+this.year;//fix this in the next 96 years...
		if(i_yrlength<1||i_yrlength==3||this.year<cY)this.year=-1;
	}
}
var o_caldiv=0;var calopen=0;
var t_calcloser = null;
function buildDate(s_monthdate){
	a_monthdate = s_monthdate.split("|");
	bd_oMonth = document.getElementById(a_monthdate[0]);
	bd_oDate = document.getElementById(a_monthdate[1]);
	i_month = bd_oMonth.selectedIndex+1;
	i_date = bd_oDate.selectedIndex+1;
	s_date = "";
	s_date = i_month+"/"+i_date+"/";
	s_date+= (i_month < cM) ? cY+1 : cY;
	return s_date;
}
function makeCalendar(v_input,s_from,s_to,s_mode){
	o_input = v_input;
	if(isPCIE){
		document.getElementById("calbox").innerHTML="<iframe id=\"calframe\" src=\"javascript:'calendar'\" scrolling=\"no\" marginheight=\"0\" marginwidth=\"0\" frameborder=\"0\"></iframe>"
		s_iecalcss = "<link rel='STYLESHEET' type='text/css' href='"+document.getElementById("calendarcss").href+"' />";
		o_califrame = document.getElementById("calframe")
		top.calframe.document.open();
		top.calframe.document.write("<html><head>"+s_iecalcss+"</head><body id='calbox' class='calendar'></body></html>");
		top.calframe.document.close();	
	}
	i_numcal = 0;
	if(isPCIE)document.getElementById("calframe").className="calframe";
	o_calbody = (isPCIE) ? top.calframe.document.getElementById("calbox") : document.getElementById("calbox");
	o_calbody.innerHTML="";
	o_udate = (typeof(o_input)=="object") ? new splitDate(o_input.value,'us') : new splitDate(buildDate(o_input),'us') ;
	a_from = s_from.split("|");
	v_from = (a_from.length==1) ? new splitDate(document.getElementById(s_from).value,s_mode) : new splitDate(buildDate(s_from),s_mode);
	a_to = s_to.split("|");
	v_to = (a_to.length==2) ? new splitDate(buildDate(s_to),s_mode) : (document.getElementById(s_to)) ? new splitDate(document.getElementById(s_to).value,s_mode) : new splitDate("",s_mode);
	d_makefrom = (v_from.month!=-1) ? new Date(v_from.year,v_from.month,v_from.date) : 0;
	d_maketo = (v_to.month!=-1) ? new Date(v_to.year,v_to.month,v_to.date) : 0;
	if(o_udate.month!=-1&&o_udate.year!=-1&&o_udate.date!=-1){
		newY=o_udate.year;
		newM=o_udate.month;
		newD=o_udate.date;
		vm_makeCal(o_udate.month);	
	}else if(o_udate!=document.getElementById(s_from)&&d_makefrom){
		newY=v_from.year;
		newM=v_from.month;
		newD=v_from.date;
		vm_makeCal(v_from.month);	
	}else if(o_udate!=document.getElementById(s_to)&&d_maketo){
		newY=v_to.year;
		newM=v_to.month;
		newD=v_to.date;
		vm_makeCal(v_to.month);	
	}else{
		newY=cY;newM=cM;newD=cD;
		vm_makeCal(cM);	
	}
}
function openCal(v_input,s_from,s_to,s_cal,s_parent,s_mode,s_caldir){
	var obj_parent;
	var prior_o_parent;
	var parent_offset      = 0;
	var parent_offset_left = 0;
	if(o_parent) o_parent.className = "cbrow";//(s_inputtype=="object") ? "cbcalrow" : "cbrow"; 	
	clearTimeout(t_calcloser);
	s_inputtype = typeof(v_input);
	a_v_input = null;
	if(s_inputtype!="object") a_v_input = v_input.split("|");
	calopen = 0;
	o_caldiv = top.document.getElementById(s_cal);	
	o_caldiv.style.display = "block";
	o_caldiv.className = "calboxon";
	makeCalendar(v_input,s_from,s_to,s_mode)
	//get objects
	o_inputright = (s_inputtype!="object") ?  document.getElementById(a_v_input[1]) : v_input;
	o_parent = document.getElementById(s_parent);
	o_cal = document.getElementById(s_cal);
	o_parent.className+=" cbrowon";

	while (o_parent.tagName != "BODY") {
		parent_offset = parent_offset + o_parent.offsetTop;
		parent_offset_left = parent_offset_left + o_parent.offsetLeft;
		prior_o_parent = o_parent;
		o_parent = o_parent.offsetParent;
	}
	
	// set it back to the parent prior tot he body
	o_parent = prior_o_parent;
	
	i_calx = (parseInt(document.body.clientWidth) - parseInt(o_parent.clientWidth)) / 2;
	i_calx = parseInt(o_inputright.offsetLeft) + parseInt(parent_offset_left) + parseInt(o_inputright.offsetWidth); 
	i_caly = parseInt(parent_offset) - (parseInt(o_cal.offsetHeight)/2);

	if (i_caly < 0){
	i_caly = (i_caly * -1) + 30;
	}
	o_cal.style.top = (i_caly>0) ? i_caly+"px" : "0px";
	o_cal.style.left = i_calx+"px";
	setTimeout("calopen = 1",500);
}
function closeCal(){
	if(o_caldiv&&calopen)t_calcloser = setTimeout("hideCalendar()",500);
	calopen=0;
}
window.onclick=closeCal;
window.document.onclick=closeCal;

