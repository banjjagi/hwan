;������ȣ �Է�
Loop, Read, D:\db\nong.txt
{
	CoordMode,mouse,relative
	LineNumber=%A_Index%
	
	Loop, parse, A_LoopReadLine, csv
	{
		SetTitleMatchMode, 2
		WinActivate, �ǰ����������ý��ۡ�������
		sleep, 1000
		click, left, 150,150
		click, left, 150,150
		sleep, 2000
		send,%a_loopfield%
		sleep,1000
		send,{enter}
		sleep, 3000
		click, left, 150,550
		click, left, 150,550
		winwait, �˸�
		sleep, 1000
		WINCLOSE, �˸�
		fileappend,%clipboard%`n, D:\db\nong_rrno_ins.csv
	}
}	
F2::Pause
F3::ExitApp