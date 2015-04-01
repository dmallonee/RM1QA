var cF=document.all.CalFrame;var cW=window.frames.CalFrame;var g_tid=0;var g_cP,g_eD,g_eDP,g_dmin,g_dmax,g_htm;
function CB(){event.cancelBubble=true}
function SCal(cP,eD,eDP,dmin,dmax,htm){clearTimeout(g_tid);var s=(g_eD==eD);g_cP=cP;g_eD=eD;g_eDP=eDP;g_dmin=dmin;g_dmax=dmax;g_htm=htm;WaitCal(true,s);}
function CancelCal(){clearTimeout(g_tid);cF.style.display="none";}
function WaitCal(i,s)
{
	if(null==cW.g_fCL||false==cW.g_fCL)
	{
		if(i){if(s&&"block"==cF.style.display){cF.style.display="none";return;}PosCal(g_cP);cF.style.display="block";}
		g_tid=setTimeout("WaitCal()", 200);
	}
	else cW.DoCal(g_cP,g_eD,g_eDP,g_dmin,g_dmax);
}

function PosCal(cP)
{
	var dB=document.body;var eL=0;var eT=0;
	for(var p=cP;p&&p.tagName!='BODY';p=p.offsetParent){eL+=p.offsetLeft;eT+=p.offsetTop;}
	var eH=cP.offsetHeight;var dH=cF.style.pixelHeight;var sT=dB.scrollTop;
	if(eT-dH>=sT&&eT+eH+dH>dB.clientHeight+sT)eT-=dH;else eT+=eH;
	cF.style.left=eL;cF.style.top=eT;
}
function GetDowStart() {return 0;}
function GetDateFmt() {return "mmddyy";}
function GetDateSep() {return "/";}
function ShowCalendar(eP,eD,eDP,dmin,dmax){
	var htm="cal.htm";
	SCal(eP,eD,eDP,dmin,dmax,htm);
	}