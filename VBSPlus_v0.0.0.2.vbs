' VBS+ Project
' Aim to build a more advanced VBScripting experience by improving the vbs grammar
' Now I'm trying to reconstruct the original VBS grammar first

' Known Issues
' 1. In-Line Comments without space split like "StatementsREMcomments", cannot be removed


Dim V: Set V = New VBSPlusNameSpace

Class VBSPlusNameSpace
	Private FSO, ws, SA, ADO, wn
	Private SelfFolderPath, UserName, Self, IDLECode
	Private VBSInterpreter, VJson
	Public Developer
	Public NameSpace_Version

	Private Sub Class_Initialize()
		Const Developer = True
		Const NameSpace_Version = "0.0.0.1"
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		Set ws = CreateObject("Wscript.Shell")
		Set SA = CreateObject("Shell.Application")
		'Set ADO = CreateObject("ADODB.STREAM")
		'Set wn = CreateObject("Wscript.Network")

		Call GetUAC(1, False)

		SelfFolderPath = FormatPath(FSO.GetFile(WScript.ScriptFullName).ParentFolder.Path)
		'UserName = wn.UserName
		'Self = FSO.OpenTextFile(Wscript.ScriptFullName).ReadAll
		
		Set VBSInterpreter = New VBSPlus_Interpreter
		Set VJson = New VbsJson
		
		' ## Debugging ##
		VBSInterpreter.NoComments = True
		'MsgBox VBSInterpreter.GetCodeWithLayer(Self, InStr(1, Self, "Public Fun" & "ction GetU"))
		
		Dim DebugPath:   DebugPath   = SelfFolderPath & "Debug Outputs\"
		Dim DebugScript: DebugScript = "F:\VBSShell\VBShell_v1.0.0.7.vbs"
		' "F:\VBSShell\VBShell_v1.0.0.7.vbs"
		' "F:\VBS+ Project\变态的If语句.vbs"
		' Wscript.ScriptFullName
		
		If Not FSO.FolderExists(DebugPath) Then FSO.CreateFolder(DebugPath)
		
		' Load Script
		VBSInterpreter.LoadScript(DebugScript)
		
		' Output logical lines
		FSO.CreateTextFile(DebugPath & "LogicalLines.txt", True).Write VJson.Encode(VBSInterpreter.ScriptLogicalLines)
		ws.Run "notepad.exe """ & DebugPath & "LogicalLines.txt" & """", 1, True
		
		' Output words
		FSO.CreateTextFile(DebugPath & "AllWords.txt", True).Write VJson.Encode(VBSInterpreter.ScriptWords)
		ws.Run "notepad.exe """ & DebugPath & "AllWords.txt" & """", 1, True
		
		' Compress code into one line
		If VBSInterpreter.NoComments Then
			Dim OneLine
			OneLine = Join(VBSInterpreter.ScriptLogicalLines, ":")
			FSO.CreateTextFile(DebugPath & "OneLine.txt", True).Write OneLine
			ws.Run "notepad.exe """ & DebugPath & "OneLine.txt" & """", 1, True
		End If
	End Sub
	
	Property Get Version()
		Version = NameSpace_Version
	End Property

	Public Function GetUAC(ByVal Host, ByVal Hide)
		''' GetUAC By PY-DNG; Version 1.7 '''
		' 最近更新：更换了UAC判断方式，不再占用命令行参数，兼容了没有UAC机制的更老版本Windows系统（如XP，2003）；简化了代码的表示
		On Error Resume Next: Err.Clear
		Dim HostName, Args, i, Argv, TFPath, HaveUAC
		If Host = 1 Then HostName = "wscript.exe"
		If Host = 2 Then HostName = "cscript.exe"
		' Get All Arguments
		Set Argv = WScript.Arguments
		For Each Arg in Argv
			Args = Args & " " & Chr(34) & Arg & Chr(34)
		Next
		' Test If We Have UAC
		TFPath = FSO.GetSpecialFolder(WindowsFolder) & "\system32\UACTestFile"
		FSO.CreateTextFile TFPath, True
		HaveUAC = FSO.FileExists(TFPath) And Err.number <> 70
		If HaveUAC Then FSO.DeleteFile TFPath, True
		' If No UAC Then Get It Else Check & Correct The Host
		If Not HaveUAC Then
			SA.ShellExecute "wscript.exe", "//e:VBScript " & Chr(34) & WScript.ScriptFullName & chr(34) & Args, "", "runas", 1
			WScript.Quit
		ElseIf LCase(Right(WScript.FullName,12)) <> "\" & HostName Then
			ws.Run HostName & " //e:VBScript """ & WScript.ScriptFullName & """" & Args, Int(Hide)+1, False
			WScript.Quit
		End If
		If Host = 2 Then ExecuteGlobal "Dim SI, SO: Set SI = Wscript.StdIn: Set SO = Wscript.StdOut"
	End Function

	Public Function FormatPath(ByVal Path)
		If Not Right(Path, 1) = "\" Then
			Path = Path & "\"
		End If
		FormatPath = Path
	End Function

	Public Function CreateTempPath(ByVal IsFolder)
		Dim TempPath
		TempPath = FSO.GetSpecialFolder(2) & "\" & FSO.GetTempName()
		If IsFolder Then TempPath = FormatPath(TempPath)
		CreateTempPath = TempPath
	End Function
	
	Public Function Import(ByVal ModelPath)
		'
	End Function
End Class

Class VBSPlus_Interpreter
	Private Sub Class_Initialize()
		Developer = True
		Interpreter_Version = "0.0.0.2"
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		Blank         = " " & vbTab & vbCr & vbLf
		Whitespace    = " " & vbTab
		ConnectChars  = "+-*/\^,&<>=("
		SplitChars    = "+-*/\^,&<>=()[] " & vbTab
		ExpConnecters = Array("mod","is","not","and","or","xor","eqv","imp")
	End Sub
	
	Private Sub Class_Terminate()
		Set FSO = Nothing
	End Sub
	
	Private FSO
	Private Whitespace, Blank, ExpConnecters, ConnectChars, SplitChars
	Private ScriptPath, ScriptShortPath
	Public Developer, ClearREM
	Public Interpreter_Version
	
	Property Let NoComments(ByVal Bool)
		ClearREM = Bool
	End Property
	
	Property Get NoComments()
		NoComments = ClearREM
	End Property
	
	Property Get Version()
		Version = Interpreter_Version
	End Property
	
	Property Get ScriptLogicalLines()
		ScriptLogicalLines = LogicalLines
	End Property
	
	Property Get ScriptWords()
		ScriptWords = AllWords
	End Property
	
	Property Get ScriptFullPath()
		ScriptFullPath = ScriptPath
	End Property
	
	Property Get ScriptCode()
		ScriptCode = Code
	End Property
	
	Public Function LoadCode(ByVal AllCode)
		LoadCode = GetLogicalLines(AllCode)
	End Function
	
	Public Function LoadScript(ByVal Path)
		' Deal Arguements
		If FSO.FolderExists(Path)   Then Call PopupDebugInfo("LoadScript Error - Path is a Folder", "VBSPlus Interpreter", "Error", 0, 0)
		If Not FSO.FileExists(Path) Then Call PopupDebugInfo("LoadScript Error - File Not Found"  , "VBSPlus Interpreter", "Error", 0, 0)
		
		' Read Script
		Dim AllCode: AllCode = FSO.OpenTextFile(Path).ReadAll
		
		' Deal Script
		Call GetLogicalLines(AllCode)
		ScriptShortPath = FSO.GetFile(Path).ShortPath
		LoadScript      = LogicalLines
	End Function
	
	Public Function GetCodeWithLayer(ByVal CodeAll, ByVal Index)
		Dim Layer: Layer = 0                 ' Layer splited by brackets
		Dim LayerReturn: LayerReturn = False ' Flag wheather Layer has increased
		Dim InStrDbl: InStrDbl = False       ' Flag wheather the current Char is in a double quotated string
		Dim InREM: InREM = False             ' Flag wheather the current Char is part of an annotation
		For i = Index To Len(CodeAll)
			Char = Mid(CodeAll, i, 1)
			Select Case Char
				Case "(", "[", "{"
					If Not InStrDbl Then Layer = Layer + 1
				Case ")", "]", "}"
					If Not InStrDbl Then Layer = Layer - 1
				Case """"
					InStrDbl = Not InStrDbl
				Case "'"
					If Not InStrDbl Then InREM = True
				Case vbCr, vbLf
					InREM    = False
					InStrDbl = False
			End Select
			If Layer = 0 Then
				If LayerReturn Then
					GetCodeWithLayer = Mid(CodeAll, Index, i-Index+1)
					Exit Function
				End If
			Else
				If Not LayerReturn Then LayerReturn = True
			End If
		Next
	End Function
	
	Private Code                ' Code we inpterprets
	Private Char                ' Current Char
	Private CharIndex           ' Index of Last Not-Empty Char
	Private i                   ' For...Next Index
	Private InVarName           ' Flag wheather the current Char is in a quotated variant name like "[If Do For]"
	Private InStrDbl            ' Flag wheather the current Char is in a double quotated string
	Private InREM               ' Flag wheather the current Char is part of an annotation
	Private IfCompressed        ' Flag wheather the current If expression(if exist) is compressed
	Private IfLayer             ' Layer of If Expression
	Private CpIfLayer           ' Layer of compressed If Expression (actually we don't know wheather this line is in a compressed if expression when we haven't finished reading this line yet, so this Flag actually stores the layer of If expression in current physical-line)
	Private IfDefLine           ' Flag wheather the current PHYSICAL-LINE defines an If expression
	Private InCaseVar           ' Status wheather the current word is in a case variant
	Private LineContinue        ' Status wheather the logical line needs to be connected into last line because of last line's "_": -1=No, 0=NextLine, 1=ThisLine, 2=ThisLine & NextLine
	Private Line                ' Current Line String
	Private LineStart           ' Current Line initial index
	Private LogicalLines()      ' All Lines
	Private LineUBound          ' Current Line's quote number
	Private OriWord             ' Current Word without lowercasing
	Private Word                ' Current Word
	Private LastWord            ' Last Word/Current Word's adj.
	Private ScndWord            ' Second Previous Word/Last Word's adj 
	Private WordLength          ' The length of the word we finding now
	Private WordStart           ' The index where current word starts
	Private AllWords()          ' All Words Array
	Private WordsUBound         ' Equals UBound(Words) - 1; Always prepares for the next Word
	Private LineWords()         ' Words Array of current Line
	Private LWordsUbound        ' Equals UBound(LineWords) - 1; Always prepares for the next Word
	
	Private Function InitVars(AllCode)
		Code         = AllCode
		InVarName    = False
		InStrDbl     = False
		InREM        = False
		IfCompressed = False
		IfLayer      = 0
		CpIfLayer    = 0
		IfDefLine    = False
		InCaseVar    = -1
		LineContinue = -1
		Line         = ""
		LineStart    = 1
		ReDim LogicalLines(0)
		LineUBound   = 0
		Word         = ""
		LastWord     = ""
		ScndWord     = ""
		WordLength   = 0
		WordStart    = 1
		ReDim AllWords(0)
		WordsUBound  = 0
	End Function
	
	Public Function GetLogicalLines(ByVal AllCode)
		' Init variants
		Call InitVars(AllCode)
		' Add vbCrLf to the end of code: make sure we will not miss the last line
		Code = Code & vbCrLf
		
		' Scan Lines
		For i = 1 To Len(Code)
			Char = Mid(Code, i, 1)
			If Char <> " " And Char <> vbTab And Char <> vbCr And Char <> vbLf And Not(InVarName) And Not(InStrDbl) And Not(InREM) Then CharIndex = i
			Select Case Char
				Case "["
					If Not(InStrDbl)  And Not(InREM)   Then InVarName = True
				Case "]"
					If Not(InStrDbl)  And Not(InREM)   Then InVarName = False
				Case """"
					If Not(InVarName) And Not(InREM)   Then InStrDbl  = Not  InStrDbl
				'Case "'"
				'	If Not(InVarName) And Not(InStrDbl) And Not(InREM)  Then Call NextWord(False): InREM = True
				Case " ", vbTab, ":", "'"
					If Not(InVarName) And Not(InStrDbl) And Not(InREM)  Then Call NextWord(False)
				Case vbCr, vbLf
					' If InStrDbl Flag = True then occurs an error
					If InStrDbl Then Call PopupDebugInfo("Unfinished String", "VBSPlus Interpreter", "Error", LineUBound, i): WScript.Quit
					' No hurry for if detection, deal with word first
					Call NextWord(True)
					' Start a new Line
					Call NextLine(True)
					' If LastWord = "then" or if not defined in this physical-line, it means that it's a NOT-COMPRESSED If expression
					If CpIfLayer = 0 Or Not(IfDefLine) Or LastWord = "then" Or LineContinue > -1 Then IfCompressed = False Else IfCompressed = True
					' Physical-Line ends, reset IfDefLine to False
					If LineContinue <= 0 Then IfDefLine = False
					' If IfLayer Layer > 1 then occurs an error
					If IfCompressed And CpIfLayer > 1 Then Call PopupDebugInfo("Unfinished If Expression", "VBSPlus Interpreter", "Error", LineUBound, i): WScript.Quit
					' Deal with compressed If expression
					' This happens because compressed If expression in a natural line may ends without the last "End If" when a new line is splited with vbCr or vbLf
					If IfCompressed And CpIfLayer = 1 And LineContinue <= 0 Then
						' Append "End If" into LogicalLines
						ReDim Preserve LogicalLines(LineUBound)
						LogicalLines(LineUBound) = "End If"
						LineUBound = LineUBound + 1
						IfLayer = IfLayer - 1
					End If
					If LineContinue <= 0 Then CpIfLayer = 0
			End Select
		Next
		GetLogicalLines = LogicalLines
	End Function
	
	Private Function NextWord(ByVal AtEndOfLine)
		' If Word is in comments and we don't care about comments
		If ClearREM And InREM Then
			' Mark new word start index
			WordStart = SkipBlank(Code, i+1)
			Exit Function
		End If
		
		' If Word dealed before
		If WordStart >= i Then 
			If Char = "'" Then InREM = True
			If Char = ":" Then NextLine(AtEndOfLine)
			Exit Function
		End If
		
		' Current Word
		OriWord = Mid(Code, WordStart, i-WordStart)
		Word = LCase(OriWord)
		
		' "REM" in Word's beginning means comments starts form here
		' Throw the whole word into trash
		If Left(Word, 3) = "rem" Then
			InREM = True
			' Mark new word start index
			WordStart = SkipBlank(Code, i+1)
			Exit Function
		End If
		
		' ## Save to Words Array ##
		' AllWords array
		ReDim Preserve AllWords(WordsUBound)
		AllWords(WordsUBound) = OriWord
		WordsUBound = WordsUBound + 1
		
		' LineWords Array
		ReDim Preserve LineWords(LWordsUbound)
		LineWords(LWordsUbound) = OriWord
		LWordsUbound = LWordsUbound + 1
		' ## Save to Words Array ##
		
		' Layer IfLayer In/Decrease and also split a new line
		If Word = "then" Then Call NextLine(AtEndOfLine)
		If Word = "else" And LastWord <> "case" Then Call NextLine(AtEndOfLine)
		
		' Expand Select Case block
		If InCaseVar = 0 Then 
			' Assume the case variant has ended
			Call NextLine(AtEndOfLine)
			InCaseVar = 1
		ElseIf InCaseVar = 1 Then
			' Check wheather there is a mistake(last Line with case has actually NOT ended, but we treated it as it has already ended at last Word)
			If InStr(1, ConnectChars, Left(Word, 1)) > 0 Or _
				ItemInArray(ExpConnecters, Word) Or _
				InStr(1, ConnectChars, Right(LastWord, 1)) > 0 Or _
				ItemInArray(ExpConnecters, LastWord) > 0 Then
				' If we found an operator at this Word's beginning, it means tha we had made a mistake
				' We ASSUME the case variable has ended at last Word, but now it seems it's not
				' For remedy, we have to connect this word into last Line
				LogicalLines(LineUBound-1) = LogicalLines(LineUBound-1) & " " & Word
				LineStart = SkipWhitespace(Code, i+1)
				Erase LineWords
				LWordsUbound = 0
			Else
				InCaseVar = -1
			End If
		ElseIf InCaseVar = 1 And AtEndOfLine Then
			AtEndOfLine = -1
		End If
		If Word = "case" And LastWord <> "select" Then
			If InCaseVar = -1 Then
				InCaseVar = 0
			Else
				Call PopupDebugInfo("Case Variable Not Found", "VBSPlus Interpreter", "Error", LineUBound, i)
				WScript.Quit
			End If
		End If
		
		' LineSplit Dealing
		If Char = ":" Then Call NextLine(AtEndOfLine)
		
		' Comments Dealing
		If Char = "'" Then InREM = True
		
		' If expression layer counting & End If split new line
		If Word = "if" Then
			If LastWord = "end" Then
				Call NextLine(AtEndOfLine)
				IfLayer = IfLayer - 1
				If IfDefLine And CpIfLayer > 0 Then CpIfLayer = CpIfLayer - 1
			Else
				IfDefLine = True
				IfLayer   = IfLayer   + 1
				CpIfLayer = CpIfLayer + 1
			End If
		End If
		
		' Save to recent Words
		ScndWord = LastWord
		LastWord = Word
		
		' Clear Word
		Word = ""
		
		' Mark new word start index
		WordStart = SkipBlank(Code, i+1)
	End Function
	
	Private Function NextLine(ByVal AtEndOfLine)
		'On Error Resume Next: Err.Clear
		''' Starts a new line and deal with flags and line indexs'''
		' LineContinue Dealing
		If AtEndOfLine And Mid(Code, CharIndex, 1) = "_" And i > 1 Then
			If InStr(1, SplitChars, Mid(Code, CharIndex-1, 1)) > 0 Then
				' Mark LineContinue
				LineContinue = LineContinue + 1
				' Correct Word
				If LastWord <> "_" Then
					' Remove "_" from last Word
					LineWords(LWordsUbound-1) = Left(LineWords(LWordsUbound-1), Len(LineWords(LWordsUbound-1))-1)
					AllWords(WordsUBound-1) = Left(AllWords(WordsUBound-1), Len(AllWords(WordsUBound-1))-1)
				Else
					' Remove last Word
					LWordsUbound = LWordsUbound - 1: ReDim Preserve LineWords(LWordsUbound - 1)
					 WordsUBound =  WordsUBound - 1: ReDim Preserve  AllWords( WordsUBound - 1)
					
					' Recover LastWord and ScndWord
					LastWord = LCase(AllWords(WordsUBound-1))
					ScndWord = LCase(AllWords(WordsUBound-2))
					
					' Recover CpIfLayer
					' CpIfLayer can not be decreased in NextWord() because NextWord() is called before NextLine() in Char Case (vbCr, vbLf)
					If LastWord = "if" And ScndWord = "end" Then CpIfLayer = CpIfLayer - 1
				End If
			End If
		End If
		
		' New Line
		If i > LineStart And LWordsUbound > 0 Then
			' Deal this Line
			' To format code into a more standarized one, we don't just use Mid(Code) but Join(LineWords) for a Line
			'Line = Mid(Code, LineStart, i-LineStart) ' This one works as well, but we don't use this
			Line = Join(LineWords, " ")
			Erase LineWords
			LWordsUbound = 0
			If  LineContinue >= 1 Then
				' LineContinue Dealing
				' If this Line is the continue of last Line, DO NOT save as a new logical Line, just add to last Line
				LogicalLines(LineUBound-1) = LogicalLines(LineUBound-1) & " " & Line
				LineContinue = LineContinue - 2 ' This Line(1) to no-continue(-1), this and next Line(2) to next Line(0)
			Else
				' Save current Line to a new logical Line
				ReDim Preserve LogicalLines(LineUBound)
				LogicalLines(LineUBound) = Line ' i-LineStart does NOT contains vbCr, vbLf, ":", and any char at this place
				' Prepare for next Line
				LineUBound = LineUBound + 1
			End If
			If LineContinue = 0 Then LineContinue = 1
			' If there is an Else in the end of this logical line, we need it to be stored in a separate line
			If Word = "else" And Not(LastWord = "case") And LCase(LogicalLines(LineUBound-1)) <> "else" Then
				LogicalLines(LineUBound-1) = Left(LogicalLines(LineUBound-1), Len(LogicalLines(LineUBound-1)) - 5) ' -5 not -4 aims to del the " "
				ReDim Preserve LogicalLines(LineUBound)
				LogicalLines(LineUBound) = "Else"
				LineUBound = LineUBound + 1
			End If
			' If there is an End If in the end of this logical line, we need it to be stored in a separate line
			If Word = "if" And LastWord = "end" And LCase(LogicalLines(LineUBound-1)) <> "end if" Then
				LogicalLines(LineUBound-1) = Left(LogicalLines(LineUBound-1), Len(LogicalLines(LineUBound-1)) - 7) ' -7 not -6 aims to del the " "
				ReDim Preserve LogicalLines(LineUBound)
				LogicalLines(LineUBound) = "End If"
				LineUBound = LineUBound + 1
			End If
			'填充模板：" = " & CStr() & vbCrLf & vbTab & _
			'PopupDebugInfo _
			'	"LL(L-1) = """ & LogicalLines(LineUBound-1) & """" & vbCrLf & vbTab & _
			'	"Line = """ & Line & """" & vbCrLf & vbTab & _
			'	"LineContinue = " & CStr(LineContinue) & vbCrLf & vbTab & _
			'	"CpIfLayer = " & CStr(CpIfLayer) & vbCrLf & vbTab & _
			'	"CharIndex = " & CStr(CharIndex) & vbCrLf & vbTab & _
			'	"Mid(Code, CharIndex, 1) = " & CStr(Mid(Code, CharIndex, 1)) & vbCrLf & vbTab & _
			'	"Word = """ & Word & """" & vbCrLf & vbTab & _
			'	"LastWord = """ & LastWord & """" & vbCrLf & vbTab & _
			'	"ScndWord = """ & ScndWord & """" & vbCrLf & vbTab & _
			'	"LWordsUbound = " & CStr(LWordsUbound) & vbCrLf & vbTab & _
			'	"LWordsUbound = " & CStr(WordsUbound), _
			'	"", "Debug", LineUBound, i
		End If
		
		' Move index
		If AtEndOfLine Then LineStart = SkipBlank(Code,i+1) Else LineStart = SkipWhitespace(Code, i+1) End If
		i = LineStart - 1
		' Reset Variables at end of each line
		If AtEndOfLine Then
			InREM = False
			If LineContinue <> 0 And LineContinue <> 2 Then
				InCaseVar = -1
			End If
		End If
	End Function
	
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	Public Function PopupDebugInfo(ByVal Text, ByVal Source, ByVal PopupLevel, ByVal LineNum, ByVal IndexNum)
		'If LineNum <= 420 Then Exit Function
		'If PopupLevel <> "Error" Then Exit Function
		
		Dim ShowIcon: ShowIcon = 4096
		Dim AfterFix: AfterFix = ""
		Dim ReturnNum, Path
		
		Select Case PopupLevel
			Case "Info"     : ShowIcon = ShowIcon + 64
			Case "Warning"  : ShowIcon = ShowIcon + 48
			Case "Question" : ShowIcon = ShowIcon + 32
			Case "Debug"    : ShowIcon = ShowIcon + 32
			Case "Error"    : ShowIcon = ShowIcon + 16
			Case "None"     : ShowIcon = ShowIcon + 0
		End Select
		If ScriptPath = ""      Then Path = WScript.ScriptFullName Else Path = ScriptPath
		If Source     = ""      Then Source = "VBSPlus Interpreter Ver_" & Interpreter_Version
		If PopupLevel = "Debug" Then ShowIcon = ShowIcon + vbYesNo: AfterFix = vbCrLf & vbCrLf & "Continue Debugging?"
		
		ReturnNum = MsgBox( _
			"Script:	" & Path     & vbCrLf & _
			"Line:	"     & LineNum  & vbCrLf & _
			"Index:	"     & IndexNum & vbCrLf & _
			"Text: 	"     & Text     & vbCrLf & _
			"Source:	" & Source   & AfterFix _
			, ShowIcon, "VBSPlus Interpreter"   _
		)
		If ReturnNum <> vbOK And ReturnNum <> vbYes Then WScript.Quit
	End Function
	
	Private Function SkipWhitespace(ByRef str, ByVal idx)
		''' Skip all Whitespaces and vbTabs; Use for preventing CrLf-Leak '''
		Do While idx <= Len(str) And _
			InStr(Whitespace, Mid(str, idx, 1)) > 0
			idx = idx + 1
		Loop
		SkipWhitespace = idx
	End Function
	
	Private Function SkipBlank(ByRef str, ByVal idx)
		''' Skip all Whitespaces, vbTabs, vbCrs and vbLfs '''
		Do While idx <= Len(str) And _
			InStr(Blank, Mid(str, idx, 1)) > 0
			idx = idx + 1
		Loop
		SkipBlank = idx
	End Function
	
	Private Function ItemInArray(ByRef Arr, ByRef Item)
		ItemInArray = False
		Dim i
		For i = 1 To UBound(Arr)
			If Arr(i) = Item Then
				ItemInArray = True
				Exit Function
			End If
		Next
	End Function
End Class




Class VbsJson
	'Author: Demon
	'Date: 2012/5/3
	'Website: http://demon.tw
	Private Whitespace, NumberRegex, StringChunk
	Private b, f, r, n, t
	
	Private Sub Class_Initialize
		Whitespace = " " & vbTab & vbCr & vbLf
		b = ChrW(8)
		f = vbFormFeed
		r = vbCr
		n = vbLf
		t = vbTab
		
		Set NumberRegex = New RegExp
		NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
		NumberRegex.Global = False
		NumberRegex.MultiLine = True
		NumberRegex.IgnoreCase = True
		
		Set StringChunk = New RegExp
		StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
		StringChunk.Global = False
		StringChunk.MultiLine = True
		StringChunk.IgnoreCase = True
	End Sub
	
	'Return a JSON string representation of a VBScript data structure
	'Supports the following objects and types
	'+-------------------+---------------+
	'| VBScript          | JSON          |
	'+===================+===============+
	'| Dictionary        | object        |
	'+-------------------+---------------+
	'| Array             | array         |
	'+-------------------+---------------+
	'| String            | string        |
	'+-------------------+---------------+
	'| Number            | number        |
	'+-------------------+---------------+
	'| True              | true          |
	'+-------------------+---------------+
	'| False             | false         |
	'+-------------------+---------------+
	'| Null              | null          |
	'+-------------------+---------------+
	Public Function Encode(ByRef obj)
		Dim buf, i, c, g
		Set buf = CreateObject("Scripting.Dictionary")
		Select Case VarType(obj)
			Case vbNull
			buf.Add buf.Count, "null"
			Case vbBoolean
			If obj Then
				buf.Add buf.Count, "true"
			Else
				buf.Add buf.Count, "false"
			End If
			Case vbInteger, vbLong, vbSingle, vbDouble
			buf.Add buf.Count, obj
			Case vbString
			buf.Add buf.Count, """"
			For i = 1 To Len(obj)
				c = Mid(obj, i, 1)
				Select Case c
					Case """" buf.Add buf.Count, "\"""
					Case "\"  buf.Add buf.Count, "\\"
					Case "/"  buf.Add buf.Count, "/"
					Case b    buf.Add buf.Count, "\b"
					Case f    buf.Add buf.Count, "\f"
					Case r    buf.Add buf.Count, "\r"
					Case n    buf.Add buf.Count, "\n"
					Case t    buf.Add buf.Count, "\t"
					Case Else
					If AscW(c) >= 0 And AscW(c) <= 31 Then
						c = Right("0" & Hex(AscW(c)), 2)
						buf.Add buf.Count, "\u00" & c
					Else
						buf.Add buf.Count, c
					End If
				End Select
			Next
			buf.Add buf.Count, """"
			Case vbArray + vbVariant
			g = True
			buf.Add buf.Count, "["
			For Each i In obj
				If g Then g = False Else buf.Add buf.Count, ","
				buf.Add buf.Count, Encode(i)
			Next
			buf.Add buf.Count, "]"
			Case vbObject
			If TypeName(obj) = "Dictionary" Then
				g = True
				buf.Add buf.Count, "{"
				For Each i In obj
					If g Then g = False Else buf.Add buf.Count, ","
					buf.Add buf.Count, """" & i & """" & ":" & Encode(obj(i))
				Next
				buf.Add buf.Count, "}"
			Else
				Err.Raise 8732,,"None dictionary object"
			End If
			Case Else
			buf.Add buf.Count, """" & CStr(obj) & """"
		End Select
		Encode = Join(buf.Items, "")
	End Function
	
	'Return the VBScript representation of ``str(``
	'Performs the following translations in decoding
	'+---------------+-------------------+
	'| JSON          | VBScript          |
	'+===============+===================+
	'| object        | Dictionary        |
	'+---------------+-------------------+
	'| array         | Array             |
	'+---------------+-------------------+
	'| string        | String            |
	'+---------------+-------------------+
	'| number        | Double            |
	'+---------------+-------------------+
	'| true          | True              |
	'+---------------+-------------------+
	'| false         | False             |
	'+---------------+-------------------+
	'| null          | Null              |
	'+---------------+-------------------+
	Public Function Decode(ByRef str)
		Dim idx
		idx = SkipWhitespace(str, 1)
		
		If Mid(str, idx, 1) = "{" Then
			Set Decode = ScanOnce(str, 1)
		Else
			Decode = ScanOnce(str, 1)
		End If
	End Function
	
	Private Function ScanOnce(ByRef str, ByRef idx)
		Dim c, ms
		
		idx = SkipWhitespace(str, idx)
		c = Mid(str, idx, 1)
		
		If c = "{" Then
			idx = idx + 1
			Set ScanOnce = ParseObject(str, idx)
			Exit Function
		ElseIf c = "[" Then
			idx = idx + 1
			ScanOnce = ParseArray(str, idx)
			Exit Function
		ElseIf c = """" Then
			idx = idx + 1
			ScanOnce = ParseString(str, idx)
			Exit Function
		ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
			idx = idx + 4
			ScanOnce = Null
			Exit Function
		ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
			idx = idx + 4
			ScanOnce = True
			Exit Function
		ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
			idx = idx + 5
			ScanOnce = False
			Exit Function
		End If
		
		Set ms = NumberRegex.Execute(Mid(str, idx))
		If ms.Count = 1 Then
			idx = idx + ms(0).Length
			ScanOnce = CDbl(ms(0))
			Exit Function
		End If
		
		Err.Raise 8732,,"No JSON object could be ScanOnced"
	End Function
	
	Private Function ParseObject(ByRef str, ByRef idx)
		Dim c, key, value
		Set ParseObject = CreateObject("Scripting.Dictionary")
		idx = SkipWhitespace(str, idx)
		c = Mid(str, idx, 1)
		
		If c = "}" Then
			idx = idx + 1
			Exit Function
		ElseIf c <> """" Then
			WScript.Echo "ParseObject: Error Out Of Loop"
			WScript.Echo "c = """ & c & """, Asc(c) = " & CStr(Asc(c))
			Err.Raise 8732,,"Expecting property name"
		End If
		
		idx = idx + 1
		
		Do
			key = ParseString(str, idx)
			
			idx = SkipWhitespace(str, idx)
			If Mid(str, idx, 1) <> ":" Then
				WScript.Echo "ParseObject: Error In Loop"
				WScript.Echo "c = """ & c & """, Asc(c) = " & CStr(Asc(c))
				Err.Raise 8732,,"Expecting : delimiter"
			End If
			
			idx = SkipWhitespace(str, idx + 1)
			If Mid(str, idx, 1) = "{" Then
				Set value = ScanOnce(str, idx)
			Else
				value = ScanOnce(str, idx)
			End If
			ParseObject.Add key, value
			
			idx = SkipWhitespace(str, idx)
			c = Mid(str, idx, 1)
			If c = "}" Then
				Exit Do
			ElseIf c <> "," Then
				WScript.Echo "ParseObject: Error In Loop"
				WScript.Echo "c = """ & c & """, Asc(c) = " & CStr(Asc(c))
				Err.Raise 8732,,"Expecting , delimiter"
			End If
			
			idx = SkipWhitespace(str, idx + 1)
			c = Mid(str, idx, 1)
			If c <> """" Then
				WScript.Echo "ParseObject: Error In Loop"
				WScript.Echo "c = """ & c & """, Asc(c) = " & CStr(Asc(c))
				Err.Raise 8732,,"Expecting property name"
			End If
			
			idx = idx + 1
		Loop
		
		idx = idx + 1
	End Function
	
	Private Function ParseArray(ByRef str, ByRef idx)
		Dim c, values, value
		Set values = CreateObject("Scripting.Dictionary")
		idx = SkipWhitespace(str, idx)
		c = Mid(str, idx, 1)
		
		If c = "]" Then
			ParseArray = values.Items
			idx = idx + 1
			Exit Function
		End If
		
		Do
			idx = SkipWhitespace(str, idx)
			If Mid(str, idx, 1) = "{" Then
				Set value = ScanOnce(str, idx)
			Else
				value = ScanOnce(str, idx)
			End If
			values.Add values.Count, value
			
			idx = SkipWhitespace(str, idx)
			c = Mid(str, idx, 1)
			If c = "]" Then
				Exit Do
			ElseIf c <> "," Then
				WScript.Echo "ParseArray: Error In Loop"
				WScript.Echo "c = """ & c & """, Asc(c) = " & CStr(Asc(c))
				Err.Raise 8732,,"Expecting , delimiter"
			End If
			
			idx = idx + 1
		Loop
		
		idx = idx + 1
		ParseArray = values.Items
	End Function
	
	Private Function ParseString(ByRef str, ByRef idx)
		Dim chunks, content, terminator, ms, esc, char
		Set chunks = CreateObject("Scripting.Dictionary")
		
		Do
			Set ms = StringChunk.Execute(Mid(str, idx))
			If ms.Count = 0 Then
				WScript.Echo "ParseString: Error In Loop"
				WScript.Echo "c = """ & c & """, Asc(c) = " & CStr(Asc(c))
				Err.Raise 8732,,"Unterminated string starting"
			End If
			
			content = ms(0).Submatches(0)
			terminator = ms(0).Submatches(1)
			If Len(content) > 0 Then
				chunks.Add chunks.Count, content
			End If
			
			idx = idx + ms(0).Length
			
			If terminator = """" Then
				Exit Do
			ElseIf terminator <> "\" Then
				WScript.Echo "ParseString: Error In Loop"
				WScript.Echo "c = """ & c & """, Asc(c) = " & CStr(Asc(c))
				Err.Raise 8732,,"Invalid control character"
			End If
			
			esc = Mid(str, idx, 1)
			
			If esc <> "u" Then
				Select Case esc
					Case """" char = """"
					Case "\"  char = "\"
					Case "/"  char = "/"
					Case "b"  char = b
					Case "f"  char = f
					Case "n"  char = n
					Case "r"  char = r
					Case "t"  char = t
					Case Else Err.Raise 8732,,"Invalid escape"
				End Select
				idx = idx + 1
			Else
				char = ChrW("&H" & Mid(str, idx + 1, 4))
				idx = idx + 5
			End If
			
			chunks.Add chunks.Count, char
		Loop
		
		ParseString = Join(chunks.Items, "")
	End Function
	
	Private Function SkipWhitespace(ByRef str, ByVal idx)
		Do While idx <= Len(str) And _
			InStr(Whitespace, Mid(str, idx, 1)) > 0
			idx = idx + 1
		Loop
		SkipWhitespace = idx
	End Function
	
End Class