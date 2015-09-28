	Const ForReading = 1
	Dim objFSO, objRead, strLine, m, n, o, dtmLastBootUpTime, SecondDiff, MinuteDiff, HourDiff, DayDiff
	Dim objShell
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objShell = CreateObject("WScript.Shell") 
	numDays = objShell.Arguments.Item(0)

	StringCMD = "cmd /c net statistics workstation > here.txt"
	objShell.Run StringCMD, 0 , True
	If objFSO.FileExists ("here.txt") Then
		Set objRead = objFSO.OpenTextFile("here.txt", ForReading)
			Do Until objRead.AtEndOfStream 
			strLine = objRead.ReadLine
				If Instr (1,strLine,"since",1) > 0 Then
					m = Instr (1,strLine,"since",1) 
					n = mid(strLine, m, len (strLine) - m + 2)
					o = Replace (n,"since ", "")
					dtmLastBootUpTime = CDate (o)
					SecondDiff = DateDiff ("s", dtmLastBootUpTime, Now)
					MinuteDiff = DateDiff ("n", dtmLastBootUpTime, Now)
					HourDiff = DateDiff ("h", dtmLastBootUpTime, Now)
					DayDiff = DateDiff ("d", dtmLastBootUpTime, Now)
				End If
			Loop
		objRead.close
		Set objRead = Nothing
		objFSO.DeleteFile ("here.txt")
	End If
	

	Set strLine = Nothing
	Set o = Nothing
	Set n = Nothing
	Set m = Nothing

	Dim TotalSecond, TotalMinute, TotalHour, TotalDay
	TotalSecond = SecondDiff
	TotalMinute = 0
	If SecondDiff > 60 AND MinuteDiff = 1 Then
		TotalSecond = SecondDiff - 60
		TotalMinute = MinuteDiff
	End If

	If MinuteDiff > 1 Then
		TotalSecond = SecondDiff - (MinuteDiff * 60)
		TotalMinute = MinuteDiff
		If TotalSecond < 1 Then
			TotalSecond = TotalSecond + 60
			TotalMinute = TotalMinute - 1
		End If
	End If

	TotalHour = 0
	If (TotalMinute > 60) Then
		n = Instr (1, (TotalMinute / 60), ".",1)
		If n <> 0 Then
			TotalHour = Mid ((TotalMinute / 60), 1, n - 1)
			TotalMinute = TotalMinute - (TotalHour * 60)
		ElseIf n = 0 Then
			TotalHour = TotalMinute / 60
			TotalMinute = TotalMinute - (TotalHour * 60)
		End If
	End If	


	TotalDay = 0
	If (TotalHour > 24) Then
		n = Instr (1, (TotalHour / 24), ".",1)
		If n <> 0 Then
			TotalDay = Mid ((TotalHour / 24), 1, n - 1)
			TotalHour = TotalHour - (TotalDay * 24)
		ElseIf n = 0 Then
			TotalDay = TotalHour / 24
			TotalHour = TotalHour - (TotalDay * 24)		
		End If
	End If

	If TotalDay > numDays Then
	objShell.run "PC Reboot Script.hta"
	End If

	Set objShell = Nothing	
	Set dtmLastBootUpTime = Nothing
	Set SecondDiff = Nothing
	Set MinuteDiff = Nothing
	Set HourDiff = Nothing
	Set DayDiff = Nothing
	Set TotalSecond = Nothing
	Set TotalMinute = Nothing
	Set TotalHour = Nothing
	Set TotalDay = Nothing
	
