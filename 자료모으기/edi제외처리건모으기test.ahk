﻿MouseClick, left,  150,  890
Sleep, 100
 
jisa := 0106,0107,0108,0109,0110,0111,0112,0113,0115,0125,0126,0127,0128,0129,0130,0131,0132,0133,0134,0137,0138,0140,0141,0142,0147,0201,0202,0203,0204,0205,0206,0207,0208,0209,0210,0211,0212,0221,0222,0223,0224,0225,0226,0227,0231,0232,0232,0233,0235,0235,0236,0237,0241,0242,0243,0251,0252,0253,0254,0261,0262,0263,0264,0301,0302,0303,0304,0305,0306,0307,0308,0308,0309,0311,0312,0312,0313,0314,0315,0316,0317,0317,0318,0319,0320,0321,0322,0324,0326,0327,0328,0329,0330,0331,0332,0333,0334,0335,0338,0339,0401,0401,0401,0402,0402,0403,0404,0404,0405,0405,0405,0406,0408,0408,0416,0416,0416,0418,0501,0502,0503,0503,0505,0505,0505,0507,0508,0508,0510,0511,0551,0552,0553,0553,0554,0554,0555,0555,0557,0558,0558,0560,0562,0563,0565,0601,0602,0603,0604,0604,0604,0605,0606,0606,0606,0610,0610,0611,0612,0651,0652,0653,0653,0654,0658,0658,0660,0662,0662,0664,0664,0666,0666,0667,0667,0668,0668,0670,0670,0671,0671,0701,0702,0703,0704,0704,0704,0705,0705,0706,0707,0707,0708,0708,0716,0716,0716,0718,0718,0719,0719,0720,0721,0722,0722,0751,0752,0752,0753,0754,0754,0755,0756,0757,0759,0759,0762,0762,0765,0765,0765,0767,0769,0771,0771,0801,0802


Loop, parse, jisa, `,
{
	
	MouseClick, left,  622,40
	Sleep, 2000
	id:=
	send, %A_LoopField% {enter}{enter} 
	sleep, 2000
	period := 2010010120100430, 2010050120100831, 2010090120101231,2011010120110430, 2011050120110831, 2011090120111231,2012010120120430, 2012050120120831, 2012090120121231,2013010120130430, 2013050120130831, 2013090120121231,2014010120140430, 2014050120140831, 2014090120141231
	loop,parse,period,',
		{
		MouseClick, left, 470, 230
		sleep, 100
		MouseClick, left,  470,  170
		Sleep, 100
		Send, {TAB}
		sleep, 100
		send, %period%
		sleep, 1000
		Send, {ENTER}
		sleep,1000
		send {ENTER}
		sleep 1000
		send {ENTER}
		sleep, 30000
		Send, {F9}
		
		if (PixelSearch, x1, y1, 0, 0, 1600, 900, C8D7E6)
		{
					
		}
		
		IfWinNotActive, ahk_class XLMAIN
		Send, {ALTDOWN}c{ALTUP}
		else 

		{
			while
			{
				winactivate, ahk_id %id%
				controlsend,
				break
			}
		}
		
		IfWinActive, ahk_class XLMAIN
		IfWinNotActive, ahk_class XLMAIN
		Send, {ALTDOWN}{F4}{ALTUP}
		Sleep, 10000

		
		}

} 