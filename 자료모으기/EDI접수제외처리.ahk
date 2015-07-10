;
jisa = 0101,0103,0104,0105,0106,0107,0108,0109,0110,0111,0112,0113,0115,0125,0126,0127,0128,0129,0130,0131,0132,0133,0134,0137,0138,0140,0141,0142,0147,0201,0202,0203,0204,0205,0206,0207,0208,0209,0210,0211,0212,0221,0222,0223,0224,0225,0226,0227,0231,0232,0233,0235,0235,0236,0237,0241,0242,0243,0251,0252,0253,0254,0261,0262,0263,0264,0301,0302,0303,0304,0305,0306,0307,0308,0309,0311,0312,0313,0314,0315,0316,0317,0318,0319,0320,0321,0322,0324,0326,0327,0328,0329,0330,0331,0332,0333,0334,0335,0401,0401,0402,0402,0403,0404,0405,0405,0406,0408,0416,0418,0501,0502,0503,0505,0507,0508,0510,0511,0551,0552,0553,0554,0555,0557,0558,0560,0562,0563,0565,0601,0602,0603,0604,0605,0606,0610,0611,0612,0651,0652,0653,0654,0658,0660,0662,0664,0666,0667,0668,0670,0671,0702,0703,0704,0705,0706,0707,0708,0716,0718,0719,0720,0721,0722,0751,0752,0753,0754,0755,0756,0757,0759,0762,0765,0767,0769,0771,0771,0801,0802,0100,0200,0220,250,0230,0240,0338,0339

Loop, parse, jisa, `,
{
; 시간체크
	CoordMode, Mouse, relative

	SetTitleMatchMode, 2
	if(a_hour=23 or a_hour<7){
		shutdown, 13
	}
	Sleep, 1000
	WinActivate, 국민건강보험 
	;EDI 접수자료처리탭누르기
	mouseclick, left, 100, 100
	sleep, 1000
	;지사변경
	MouseClick, left,  600,40
	sleep, 1000
	Winwait, 지사접속환경설정
	sendinput, %A_LoopField% 
	sleep, 500
	sendinput,{enter}
	sleep, 500
	sendinput,{enter}
	sleep, 2000
	
	MouseClick, left, 300,170
	mouseclick, left, 300,170
	Sleep, 1000
	send, 2014010120141231
	sleep, 1000			
	Send, {F5}
	sleep, 500
	ifwinexist, 알림창
	{
		continue
	}
	sleep, 6000
	SetTitleMatchMode RegEx
	send, {F9}
	sleep, 3000
	winwait, ^$
	SetTitleMatchMode, 2
	winwait, Microsoft Office 인증 마법사
	sleep, 1000
	winclose, Microsoft Office 인증 마법사 
	sleep, 1000
	winwait, bida
	sleep, 3000
	winclose, bida
	Sleep, 3000
	
	WinActivate, 국민건강보험 
	MouseClick, left, 300,170
	mouseclick, left, 300,170
	Sleep, 1000
	send, 2013010120131231
	sleep, 1000			
	Send, {F5}
	sleep, 500
	ifwinexist, 알림창
	{
		continue
	}
	sleep, 6000
	SetTitleMatchMode RegEx
	send, {F9}
	sleep, 3000
	winwait, ^$
	SetTitleMatchMode, 2
	winwait, Microsoft Office 인증 마법사
	sleep, 1000
	winclose, Microsoft Office 인증 마법사 
	sleep, 1000
	winwait, bida
	sleep, 3000
	winclose, bida
	Sleep, 3000
	
	WinActivate, 국민건강보험 	
	MouseClick, left, 300,170
	mouseclick, left, 300,170
	Sleep, 1000
	send, 2012010120121231
	sleep, 1000			
	Send, {F5}
	sleep, 500
	ifwinexist, 알림창
	{
		continue
	}
	sleep, 6000
	SetTitleMatchMode RegEx
	send, {F9}
	sleep, 3000
	winwait, ^$
	SetTitleMatchMode, 2
	winwait, Microsoft Office 인증 마법사
	sleep, 1000
	winclose, Microsoft Office 인증 마법사 
	sleep, 1000
	winwait, bida
	sleep, 3000
	winclose, bida
	Sleep, 3000

	WinActivate, 국민건강보험 
	MouseClick, left, 300,170
	mouseclick, left, 300,170
	Sleep, 1000
	send, 2011010120111231
	sleep, 1000			
	Send, {F5}
	sleep, 500
	ifwinexist, 알림창
	{
		continue
	}
	sleep, 6000
	SetTitleMatchMode RegEx
	send, {F9}
	sleep, 3000
	winwait, ^$
	SetTitleMatchMode, 2
	winwait, Microsoft Office 인증 마법사
	sleep, 1000
	winclose, Microsoft Office 인증 마법사 
	sleep, 1000
	winwait, bida
	sleep, 3000
	winclose, bida
	Sleep, 3000

	WinActivate, 국민건강보험 
	MouseClick, left, 300,170
	mouseclick, left, 300,170
	Sleep, 1000
	send, 2010010120101231
	sleep, 1000			
	Send, {F5}
	sleep, 500
	ifwinexist, 알림창
	{
		continue
	}
	sleep, 6000
	SetTitleMatchMode RegEx
	send, {F9}
	sleep, 3000
	winwait, ^$
	SetTitleMatchMode, 2
	winwait, Microsoft Office 인증 마법사
	sleep, 1000
	winclose, Microsoft Office 인증 마법사 
	sleep, 1000
	winwait, bida
	sleep, 3000
	winclose, bida
	Sleep, 3000
	
	
	WinActivate, 국민건강보험 	
	MouseClick, left, 300,170
	mouseclick, left, 300,170
	Sleep, 1000
	send, 2015010120150331
	sleep, 1000			
	Send, {F5}
	sleep, 500
	ifwinexist, 알림창
	{
		continue
	}
	sleep, 6000
	SetTitleMatchMode RegEx
	send, {F9}
	sleep, 3000
	winwait, ^$
	SetTitleMatchMode, 2
	winwait, Microsoft Office 인증 마법사
	sleep, 1000
	winclose, Microsoft Office 인증 마법사 
	sleep, 1000
	winwait, bida
	sleep, 3000
	winclose, bida
	Sleep, 3000
	
}