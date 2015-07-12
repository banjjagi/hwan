
;결정번호 입력
Loop, Read, D:\db\Unpaid2.csv
{
	CoordMode,mouse,relative
	LineNumber=%A_Index%

	if(a_hour=23 or a_hour<7)
	{
		Shutdown, 13
	}
	Sleep, 1000
	
		Loop, parse, A_LoopReadLine, csv
		{
			SetTitleMatchMode, 2
			WinActivate, 국민건강보험 
			sleep, 3000
			click, left, 400,170
			click, left, 400,170
			sleep, 1000
			send,{Tab}
			sleep, 1000
			send, %A_LoopField%
			sleep, 1000
			send, {F5}
			sleep, 1000
			
			pixelsearch, Vx,Vy,0,0,150,150,0xD9F0FF,5,fastrgb
			if errorlevel=0
			{
				sleep, 1000
				send, {enter}
				fileappend,%A_LoopField% `n, D:\db\dead.csv
			}
		}	
}	

F2::Pause
F3::ExitApp