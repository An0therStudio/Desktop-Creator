'*******************************
'*   DESKTOP CREATOR TOOL      *
'*   CREATED BY: LUKE CREE     *
'* LAST MODIFIED: 12/17/2015   *
'*******************************

DIM fso, outFile, desktop, strDirectory, objFolder, objShell, StdOut
Set fso = CreateObject("Scripting.FileSystemObject")
Set objShell = WScript.CreateObject("Wscript.Shell")
Set StdOut = WScript.StdOut
desktop = objShell.SpecialFolders("Desktop")

'*********************
'*   INPUT FILE      *
'*********************

Dim wd  'Working Directory
wd = Left(WScript.ScriptFullName,InstrRev(WScript.ScriptFullName, "\"))

Dim csv
csv = CSVArray(wd&"DC_in.csv") 'Create array from file

'****************************
'*   CREATE LOG DIRECTORY   *
'****************************
strDirectory = "C:\logs"
If (fso.FolderExists(strDirectory)) Then
	strDirectory = strDirectory & "\DC"
	If (fso.FolderExists(strDirectory) = false) Then
		Set objFolder = fso.CreateFolder(strDirectory)
	End If
Else
	Set objFolder = fso.CreateFolder(strDirectory)
	strDirectory = strDirectory & "\DC"
	Set objFolder = fso.CreateFolder(strDirectory)
End If


'****************************
'*   OUTPUT FILE HEADER     *
'****************************
Set outFile = fso.CreateTextFile(strDirectory & "\DC_log_"& month(date) & day(date) & year(date) & Hour(Time) & Minute(Time) & ".txt", True)
outFile.WriteLine("*********DESKTOP CREATOR TOOL LOG*********")
outFile.WriteLine
outFile.WriteLine

'*********************
'*   PROCESS ARRAY   *
'*********************

Dim lnk, url
Dim strName, strTarget, strIcon, strWD, strLogLine, blLnk, blUrl

For r = 1 to UBound(csv,1)
	strName = csv(r,0)
	strTarget = csv(r,1)
	strIcon = csv(r,2)
	strWD = csv(r,3)
	blLnk = False
	blUrl = False
	strLogLine = "" 'Generic error line
'*************************
'*  CHECK SHORTCUT TYPE  *
'*************************
	If (StrComp(csv(r,4),"LNK",1) = 0) Then
		blLnk = True
	ElseIf (IsNull(StrComp(csv(r,4),"LNK",1)) = True) Then
		strLogLine = "Please specify shortcut type for " & strName & " and try again." 'Must have file type error.
	ElseIf (StrComp(csv(r,4),"URL",1) = 0) Then
		blUrl = True
	Else
		strLogLine = "Please supply [URL/LNK] type for " & strName & " and try again." 'Incorrect file type error.
	End If
'*************************
'*   CREATE LNK SHORTCUT *
'*************************
	If (blLnk = True) Then
		If (IsNull(strName) = false) And (IsEmpty(strName) = false) Then 'Check shortcut name was provided.
			Set lnk = objShell.CreateShortcut(desktop & "\" & strName & ".lnk")
			lnk.TargetPath = strTarget
			lnk.Description = strDescription
			lnk.WorkingDirectory = strWD
			If (IsNull(strIcon) = false) And (IsEmpty(strIcon) = false) Then
				lnk.IconLocation = strIcon
			End If
			lnk.save
		Else
			strLogLine = "Please supply name for shortcut# " & r & " and try again." 'No shortcut name error.
		End If
		If (fso.FileExists(desktop & "\" & strName & ".lnk")) And (strLogLine = "") Then
			strLogLine = "Shortcut to " & strName & " successfully created." 'LNK success message.
		End If
'*************************
'*   CREATE URL SHORTCUT *
'*************************
	Else If (blUrl = True) Then
		Set url = objShell.CreateShortcut(desktop & "\" & strName & ".url")
		url.TargetPath = strTarget
		url.save
		If (fso.FileExists(desktop & "\" & strName & ".url")) Then
			strLogLine = "Shortcut to " & strName & " successfully created."
		End If
	End If
	End If
	outfile.WriteLine(strLogLine) 'Write result
Next


'*************************
'*   OUTPUT FILE END     *
'*************************
outFile.WriteLine
outFile.WriteLine
outFile.Write Now
outFile.Close


Function CSVArray(CSVFile)
 
	Dim comma, quote
	comma = ","
	quote = Chr(34)
 
	Dim charPos, charVal
 
	Dim cellPos, colMax, colNum
	colMax  = -1
  
	Dim cellArray(), cellComplete, cellQuoted, csvRecord
 
	Dim inCsv, inRow(), rowCount
	rowCount     = -1
  'the CSV file is opened for reading and each line is read into the array inRow
	Set inCsv    = fso.OpenTextFile(CSVFile,"1",True)
	Do While Not inCsv.AtEndOfStream
		rowCount = rowCount + 1
		Redim Preserve inRow(rowCount)
		inRow(rowCount) = inCsv.ReadLine
	Loop
	inCsv.Close
 
 'Two loops then loop through the row data. The outer loop loops through each row of the CSV file
	For r = 0 to rowCount
  
		csvRecord = inRow(r)
		colNum = -1
		charPos = 0
		cellComplete = True
	'The inner loop loops through the characters in the row to separate out the columns.    
		Do While charPos < Len(csvRecord)
	'The inner loop processing has to evaluate the characters it has read to determine its situation. The cell may be contained within quotes, essential if the cell data contains a comma, or it may just be straight data.
	'If the cellComplete flag has been set by the last process loop, the loop is set for processing a new cell
			If (cellComplete = True) Then
				colNum       = colNum + 1
				cellPos      = 0
				cellQuoted   = False
				cellComplete = False
				If colNum > colMax Then
					colMax = colNum
					Redim Preserve cellArray(rowCount,colMax)
				End If              
			End If
	'The character to be processed this loop is retrieved
			charPos = charPos + 1
			cellPos = cellPos + 1
			charVal = Mid(csvRecord, charPos, 1)
	'The process checks whether the character is a quote. If it is it needs to determine whether it is either the start of a cell, the end of a cell or a quote within the cell data.
			If (charVal = quote) Then
				If (cellPos = 1) Then
					cellQuoted = True
					charVal    = ""
				Else
	'If it’s not the start of a new cell, the next character is read to see what action to take.
					Select Case Mid(csvRecord, charPos+1, 1)
					Case quote
						charPos = charPos + 1
					Case comma
						charPos = charPos + 1
						cellComplete = True 
					End Select
				End If
	'If the cell is not flagged as being quoted and the character is a comma, then the cell is complete.
			ElseIf (charVal = comma) And (cellQuoted = False) Then
				cellComplete = True
			End If
	'If the processing has determined that the character is data, the character is added to the array cell
			If (cellComplete = False) Then
				cellArray(r,colNum) = cellArray(r,colNum)&charVal
			End If
	 
		Loop
'When the row is complete the outer loop moves on to the next row, until the file is complete
	Next
	CSVArray = cellArray
End Function