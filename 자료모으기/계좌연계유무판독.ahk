;������ȣ �Է�
Loop, Read, D:\db\jung_source.txt
{
	CoordMode,mouse,relative
	LineNumber=%A_Index%
	
	Loop, parse, A_LoopReadLine, csv
	{
		SetTitleMatchMode, 2
		WinActivate, ���� ���ް��� ����
		sleep, 1000
		click, left, 150,90
		click, left, 150,90
		sleep, 1000
		send,%a_loopfield%
		sleep,1000
		send,{enter}
		sleep, 1000
		;settimer,Search,5000
		;settimer,CloseXL,100000

		Search:
 		loop{
			pixelsearch,xx,xy,270,180,370,200,0xF2F1F1,,RGB  ;123,321��ǥ���� 0x000000�� ��Ÿ�������� ��� ã�´�.
			if ErrorLevel=0
 			{
				fileappend,%a_loopfield%`n, D:\db\linkage.csv
				goto Save
			}
		return
		Save:
			fileappend,%a_loopfield%`n, D:\db\linkage.csv
		return
		last:
			sleep,1000
		return

	}
}