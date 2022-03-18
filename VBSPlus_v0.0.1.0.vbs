Dim V: Set V = New VBSPlusNameSpace

Class VBSPlusNameSpace
	Private FSO, ws, SA, ADO, wn
	Private SelfFolderPath, UserName, Self, IDLECode
	Private VBSLineParser, VBSWordSpliter, VJson
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
		
		Set VBSLineParser = New VBSPlus_LineParser
		Set VBSWordSpliter = New VBSPlus_WordSpliter
		'Set ExpExec = New ExpressionExecutor
		Set VJson = New VbsJson
		
		' ## Debugging ##
		VBSLineParser.NoComments = True
		'MsgBox VBSLineParser.GetCodeWithLayer(Self, InStr(1, Self, "Public Fun" & "ction GetU"))
		
		Dim DebugPath:   DebugPath   = SelfFolderPath & "Debug Outputs\"
		Dim DebugScript: DebugScript = "F:\Re_VBS+\变态的If语句.vbs"
		' "F:\VBSShell\VBShell_v1.0.0.7.vbs"
		' "F:\Re_VBS+\变态的If语句.vbs"
		' Wscript.ScriptFullName
		' "F:\Re_VBS+\By-Products\If Test.vbs"
		' "F:\Re_VBS+\VBS Code Beautifuier.vbs"
		' "F:\Re_VBS+\By-Products\czjt1234_bug.vbs"
		' "F:\Re_VBS+\By-Products\Time_bug.vbs"
		' "D:\Program\vbs\Re_VBS+\VBS Code Beautifuier.vbs"
		
		If Not FSO.FolderExists(DebugPath) Then FSO.CreateFolder(DebugPath)
		
		'MsgBox ExpExec.Calculate("1+- 3*4"), 0, "Numeric Calculate Result"      ' -11
		'MsgBox ExpExec.Calculate("Not 3"), 0, "Logical Calculate Result"        ' -4
		'MsgBox ExpExec.Calculate("Not 3>=3"), 0, "Logical Calculate Result 2"   ' False
		'MsgBox ExpExec.Calculate("Not True"), 0, "Logical Calculate Result 3"   ' False
		'MsgBox ExpExec.Calculate("Not 3+4<=0"), 0, "Merge Calculate Result"     ' True
		'MsgBox ExpExec.Calculate("Not trUe+ 1"), 0, "Merge Calculate Result 2"  ' -1
		'WScript.Quit
		
		'Dim VBSWordSplitor: Set VBSWordSplitor = New VBSPlus_WordSpliter
		'VBSWordSplitor.ScriptCode = FSO.OpenTextFile(DebugScript).ReadAll
		'FSO.CreateTextFile(DebugPath & "ScriptWords.txt", True).Write VJson.Encode(VBSWordSplitor.ScriptWords)
		'ws.Run "notepad.exe """ & DebugPath & "ScriptWords.txt" & """", 1, True
		'WScript.Quit
		
		' Load Script
		VBSLineParser.LoadScript(DebugScript)
		
		' Output logical lines
		FSO.CreateTextFile(DebugPath & "LogicalLines.txt", True).Write VJson.Encode(VBSLineParser.ScriptLogicalLinesArr)
		ws.Run "notepad.exe """ & DebugPath & "LogicalLines.txt" & """", 1, True
		
		' Output words
		FSO.CreateTextFile(DebugPath & "AllWords.txt", True).Write VJson.Encode(VBSLineParser.ScriptWordsArr)
		'ws.Run "notepad.exe """ & DebugPath & "AllWords.txt" & """", 1, True
		
		' Compress code into one line
		'If VBSLineParser.NoComments Then
		'	Dim OneLine
		'	OneLine = Join(VBSLineParser.ScriptLogicalLines, ":")
		'	FSO.CreateTextFile(DebugPath & "OneLine.txt", True).Write OneLine
		'	ws.Run "notepad.exe """ & DebugPath & "OneLine.txt" & """", 1, True
		'End If
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





Class VBSPlus_WordSpliter
	Private Sub Class_Initialize()
		Debugger = False
		Call JudgeInit()
		' 保留字 内置函数 内置常量 参考 Demon's Blog: demon.tw
		' 保留字
		ReservedWord = "And As Boolean ByRef Byte ByVal Call Case Class Class_Initialize Class_Terminate Const Currency Debug Dim Do Double Each Else ElseIf Empty End EndIf Enum Eqv Error Event Exit Explicit False For Function Get Goto If Imp Implements In Integer Is Let Like Long Loop LSet Me Mod New Next Not Nothing Null On Option Optional Or ParamArray Preserve Private Property Public RaiseEvent ReDim Rem Resume RSet Select Set Shared Single Static Step Stop Sub Then To True Type TypeOf Until Variant WEnd While With Xor"
		ReservedWord = Split(UCase(ReservedWord), " ")
		' 内置函数
		BuiltInFunction = "Abs Array Asc Atn CBool CByte CCur CDate CDbl CInt CLng CSng CStr Chr Cos CreateObject Date DateAdd DateDiff DatePart DateSerial DateValue Day Escape Eval Exp Filter Fix FormatCurrency FormatDateTime FormatNumber FormatPercent GetLocale GetObject GetRef Hex Hour InStr InStrRev InputBox Int IsArray IsDate IsEmpty IsNull IsNumeric IsObject Join LBound LCase LTrim Left Len LoadPicture Log Mid Minute Month MonthName MsgBox Now Oct Randomize RGB RTrim Replace Right Rnd Round ScriptEngine ScriptEngineBuildVersion ScriptEngineMajorVersion ScriptEngineMinorVersion Second SetLocale Sgn Sin Space Split Sqr StrComp StrReverse String Tan Time TimeSerial TimeValue Timer Trim TypeName UBound UCase Unescape VarType Weekday WeekdayName Year"
		BuiltInFunction = Split(UCase(BuiltInFunction), " ")
		' 内置常量
		BuiltInConstants = "vbBlack vbRed vbGreen vbYellow vbBlue vbMagenta vbCyan vbWhite vbBinaryCompare vbTextCompare vbSunday vbMonday vbTuesday vbWednesday vbThursday vbFriday vbSaturday vbUseSystemDayOfWeek vbFirstJan1 vbFirstFourDays vbFirstFullWeek vbGeneralDate vbLongDate vbShortDate vbLongTime vbShortTime vbObjectError vbOKOnly vbOKCancel vbAbortRetryIgnore vbYesNoCancel vbYesNo vbRetryCancel vbCritical vbQuestion vbExclamation vbInformation vbDefaultButton1 vbDefaultButton2 vbDefaultButton3 vbDefaultButton4 vbApplicationModal vbSystemModal vbOK vbCancel vbAbort vbRetry vbIgnore vbYes vbNo vbCr vbCrLf vbFormFeed vbLf vbNewLine vbNullChar vbNullString vbTab vbVerticalTab vbUseDefault vbTrue vbFalse vbEmpty vbNull vbInteger vbLong vbSingle vbDouble vbCurrency vbDate vbString vbObject vbError vbBoolean vbVariant vbDataObject vbDecimal vbByte vbArray WScript Wsh"
		BuiltInConstants = Split(UCase(BuiltInConstants), " ")
		' 空白字符
		Blank      = " " & vbTab & vbCr & vbLf
		Whitespace = " " & vbTab
	End Sub
	
	Private Debugger
	Private ReservedWord, BuiltInFunction, BuiltInConstants
	
	Private Sub Class_Terminate()
		' Do Nothing
	End Sub
	
	Property Get ScriptWords()
		ScriptWords = Words
	End Property
	
	Property Let ScriptCode(ByVal NewCode)
		Call Scan(NewCode)
	End Property
	
	Property Get ScriptCode()
		ScriptCode = Code
	End Property
	
	' ### Main Scaning Part ###
	Private Code
	Private i, Char
	Private StrMode, ModeThis
	Private WordStart, WordText, Word, Words(), WordsUBound
	' Word.Type: 
	' -1.Space_vbTab_Nothing 0.vbCrLf 1.Name 2.Symbol 3.Number 4.String 5.Time_Constance 6.Comments
	' 1.0 Name; 1.01 Name with [] quoted; 1.1 Reserved Word; 1.2 Builtin Function; 1.3 Builtin Constant
	
	Private Function ScanInit(ByVal AllCode)
		Code = AllCode + vbCrLf
		i    = SkipBlank(Code, 1)
		WordStart = i
		StrMode   = -1
		WordUBound = 0
		ReDim Words(0)
	End Function
	
	Public Function Scan(ByVal AllCode)
		ScanInit AllCode
		
		For i = i To Len(Code)
			Char = Mid(Code, i, 1)
			WordText = Mid(Code, WordStart, i-WordStart+1) ' Includes the current character
			ModeThis = JudgeMode(WordText)
			
			If StrMode <> ModeThis Then
				If StrMode > -1 Then Call NextWord()
				StrMode = ModeThis
			End If
			
			If Debugger And WordsUBound > 587 Then If MsgBox("[" & WordText & "]" & vbCrLf & "WordsUBound = " & CStr(WordsUBound) & vbCrLf & "是否继续调试？", vbYesNo+4096+32, "WordText Debugging i = " & CStr(i)) <> vbYes Then WScript.Quit
		Next
		Scan = Words
	End Function
	
	Private Function NextWord()
		Dim Text: Text = Mid(Code, WordStart, i-WordStart)
		Set Word = CreateObject("Scripting.Dictionary")
		If StrMode = 1 Then StrMode = JudgeName(Text)
		Word.Add "Type",  StrMode
		Word.Add "Value", Text ' Excludes the current character
		ReDim Preserve Words(WordsUBound)
		Set Words(WordsUBound) = Word
		WordsUBound = WordsUBound + 1
		WordStart = SkipWhitespace(Code, i)
		i = WordStart - 1
		If TypeName(SO) = "TextStream" Then SO.WriteLine "[" & Word("Value") & "]" & "(" & CStr(Len(Word("Value"))) & ") " & "{" & CStr(WordsUBound) & "}"
		If Debugger And WordsUBound > 39 Then If MsgBox("[" & Word("Value") & "]" & "(" & CStr(Len(Word("Value"))) & ")" & vbCrLf & "WordsUBound = " & CStr(WordsUBound) & vbCrLf & "WordStart = " & CStr(WordStart) & vbCrLf & "Char At " & CStr(WordStart) & " Is " & Mid(Code, WordStart, 1) & vbCrLf & "WordText = [" & WordText & "]" & vbCrLf & "StrMode = " & CStr(StrMode) & vbCrLf & "ModeThis = " & CStr(ModeThis) & vbCrLf & "是否继续调试？", vbYesNo+4096+32, "Word Debugging i = " & CStr(i)) <> vbYes Then WScript.Quit
	End Function
	
	' ### Word Mode Judging Part ###
	Private RegName, RegSymb, RegNumb, RegStrg, RegTime, RegCmts
	Private Letters, Numbers, Symbols
	Private LNMixed, HexChr, NameChr, TimeChr
	
	Private Function JudgeInit()
		Set RegName = New RegExp
		Set RegSymb = New RegExp
		Set RegNumb = New RegExp
		Set RegStrg = New RegExp
		Set RegTime = New RegExp
		Set RegCmts = New RegExp
		
		RegName.Pattern = "^(([a-zA-Z0-9][a-zA-Z0-9_]*)|(\[[^\n\r]+\]))$"
		RegSymb.Pattern = "^(\^|\+|-|\*|/|\\|\(|\)|>=|<=|<>|>|<|=|\.|,|&)$"
		RegNumb.Pattern = "^\d+(\.\d+)?$"
		RegStrg.Pattern = "^""[^\r\n""]*(("""")*[^\r\n""]*)*""$"
		RegTime.Pattern = "^#[ \t]*(\d{1,4}[ \t]*[-/][ \t]*\d{1,2}[ \t]*[-/][ \t]*\d{1,2})?[ \t]*(\d{1,2}[ \t]*:[ \t]*\d{1,2}[ \t]*:[ \t]*\d{1,2})?[ \t]*#$"
		RegCmts.Pattern = "^'[^\r\n]*$"
		
		Letters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
		Numbers = "0123456789"
		LNMixed = Letters & Numbers
		HexChr  = Numbers & Left(Letters, 6)
		NameChr = LNMixed & "_"
		TimeChr = LNMixed & ":/- "
		Symbols = Array("^", "+", "-", "*", "/", "\", "(", ")", ">=", "<=", "<>", ">", "<", "=", ".", ",", "&", "_", ":")
	End Function
	
	Private Function JudgeMode(ByVal CodeStr)
		JudgeMode = -1
		'If FullRegularity(RegName, CodeStr) Then JudgeMode = 1 : Exit Function
		'If FullRegularity(RegSymb, CodeStr) Then JudgeMode = 2 : Exit Function
		'If FullRegularity(RegNumb, CodeStr) Then JudgeMode = 3 : Exit Function
		'If FullRegularity(RegStrg, CodeStr) Then JudgeMode = 4 : Exit Function
		'If FullRegularity(RegTime, CodeStr) Then JudgeMode = 5 : Exit Function
		'If FullRegularity(RegCmts, CodeStr) Then JudgeMode = 6 : Exit Function
		
		If IsCrLf(CodeStr) Then JudgeMode = 0 : Exit Function
		If IsName(CodeStr) Then JudgeMode = 1 : Exit Function
		If IsSymb(CodeStr) Then JudgeMode = 2 : Exit Function
		If IsNumb(CodeStr) Then JudgeMode = 3 : Exit Function
		If IsStrg(CodeStr) Then JudgeMode = 4 : Exit Function
		If IsTime(CodeStr) Then JudgeMode = 5 : Exit Function
		If IsCmts(CodeStr) Then JudgeMode = 6 : Exit Function
		
		If Debugger And WordsUBound > 587 Then If MsgBox("[" & CodeStr & "]" & "(" & CStr(Len(CodeStr)) & ")" & vbCrLf & "是否继续调试？", vbYesNo+4096+32, "Judge Debugging i = " & CStr(i)) <> vbYes Then WScript.Quit
	End Function
	
	Private Function IsCrLf(ByVal Str)
		IsCrLf = False
		Dim i, Char_
		For i = 1 To Len(Str)
			Char_ = Mid(Str, i, 1)
			If InStr(1, vbCrLf, Char_) = 0 Then Exit Function
		Next
		IsCrLf = True
	End Function
	
	Private Function IsName(ByVal Str)
		IsName = False
		If Len(Str) = 0 Then Exit Function
		Dim Quoted: Quoted = False
		If InStr(1, Letters, Left(Str, 1)) = 0 Then If Left(Str, 1) = "[" Then Quoted = True Else Exit Function
		If UCase(Left(Str & GetNextStr(2), 3)) = "REM" Then Exit Function
		Dim i, Char_
		For i = 2 To Len(Str)
			Char_ = Mid(Str, i, 1)
			If Quoted Then
				If Char_ = "]" And i < Len(Str) Then Exit Function
				If InStr(1, vbCrLf, Char_) <> 0 Then Exit Function
			Else
				If InStr(1, NameChr, Char_) = 0 Then Exit Function
			End If
		Next
		IsName = True
	End Function
	
	Private Function IsSymb(ByVal Str)
		IsSymb = ItemInArray(Symbols, Str)
		If Char = "&" And UCase(GetNextStr(1)) = "H" Then IsSymb = False
	End Function
	
	Private Function IsNumb(ByVal Str)
		IsNumb = False
		Dim Dot: Dot = False
		Dim IsHex: IsHex = False
		If Len(Str) = 0 Then Exit Function
		If InStr(1, Numbers, Left(Str, 1)) = 0 Then
			' Deal Hex Numbers
			If Len(Str) < 2 Then Exit Function
			If UCase(Left(Str, 2)) <> "&H" Then Exit Function
			If Right(Str, 1) = "&" Then Str = Left(Str, Len(Str)-1)
			IsHex = True
		End If
		Dim i, start
		If IsHex Then start = 3 Else start = 2
		For i = start To Len(Str)
			If IsHex Then
				If InStr(1, HexChr, Mid(Str, i, 1)) = 0 Then Exit Function
			Else
				If InStr(1, Numbers, Mid(Str, i, 1)) = 0 Then
					If Dot = False And Mid(Str, i, 1) = "." Then Dot = True Else Exit Function
				End If
			End If
		Next
		IsNumb = True
	End Function
	
	Private Function IsStrg(ByVal Str)
		IsStrg = False
		If Len(Str) < 2 Then Exit Function
		If Left(Str, 1) <> """" Or Right(Str, 1) <> """" Then Exit Function
		If Char = """" And GetNextStr(1) = """" Then Exit Function
		Dim Quotes: Quotes = 0
		Dim i, Char_
		For i = 2 To Len(Str)-1
			Char_ = Mid(Str, i, 1)
			If Char_ = """" Then
				Quotes = Quotes + 1
			Else
				If Quotes Mod 2 <> 0 Then Exit Function
				Quotes = 0
			End If
		Next
		If Quotes Mod 2 <> 0 Then Exit Function
		IsStrg = True
	End Function
	
	Private Function IsTime(ByVal Str)
		' Not Perfected
		IsTime = False
		If Len(Str) < 2 Then Exit Function
		If Left(Str, 1) <> "#" Or Right(Str, 1) <> "#" Then Exit Function
		Dim i, Char_
		For i = 1 To Len(Str)
			Char_ = Mid(Str, i, 1)
			If InStr(1, TimeChr, Char_) = 0 Then Exit Function
		Next
		IsTime = True
	End Function
	
	Private Function IsCmts(ByVal Str)
		IsCmts = False
		If InStr(1, vbCrLf, Right(Str, 1)) <> 0 Then Exit Function
		If UCase(Left(Str & GetNextStr(2), 3)) = "REM" Then IsCmts = True
		If Left(Str, 1) = "'" Then IsCmts = True
	End Function
	
	Private Function JudgeName(ByVal Name)
		JudgeName = 1.0
		Name = UCase(Name)
		If ItemInArray(ReservedWord    , Name) Then JudgeName = 1.1: Exit Function
		If ItemInArray(BuiltInFunction , Name) Then JudgeName = 1.2: Exit Function
		If ItemInArray(BuiltInConstants, Name) Then JudgeName = 1.3: Exit Function
		If Mid(Name, 1, 1) = "[" Then JudgeName = 1.01: Exit Function
	End Function
	
	' ### Supporting Functions Part ###
	' Get Str that has not been scanned yet
	' If length is longer than all rest chars, then will just return the rest chars
	Private Function GetNextStr(ByVal length)
		GetNextStr = Mid(Code, i+1, length)
	End Function
	
	Private Function FullRegularity(ByVal Reg, ByVal Str)
		Dim Matches, oMatch
		Set Matches = Reg.Execute(Str)
		For Each oMatch In Matches
			If oMatch.Value = Str Then
				FullRegularity = True
				Exit Function
			End If
		Next
	End Function
	
	Private Function ItemInArray(ByRef Arr, ByRef Item)
		ItemInArray = False
		Dim i
		For i = 0 To UBound(Arr)
			If Arr(i) = Item Then
				ItemInArray = True
				Exit Function
			End If
		Next
	End Function
	
	Private Blank, Whitespace
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
End Class




Class VBSPlus_LineParser
	''' VBSPLUS LineParser '''
	Private Sub Class_Initialize()
		Developer = True
		LineParser_Version = "0.0.0.6.0"
		
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
	
	Private FSO, LINETYPEDICT
	Private Whitespace, Blank, ExpConnecters, ConnectChars, SplitChars
	Private cScriptPath, cScriptShortPath
	Public Developer, ClearREM
	Public LineParser_Version
	
	Property Let NoComments(ByVal Bool)
		ClearREM = Bool
	End Property
	
	Property Get NoComments()
		NoComments = ClearREM
	End Property
	
	Property Get Version()
		Version = LineParser_Version
	End Property
	
	Property Get ScriptLogicalLines()
		Set ScriptLogicalLines = cLogicalLines
	End Property
	
	Property Get ScriptLogicalLinesArr()
		Dim Arr, i
		Dim Lines()
		ReDim Lines(cLogicalLines.UBound)
		Arr = cLogicalLines.ToArray()
		For i = 0 To cLogicalLines.UBound
			Lines(i) = cLogicalLines(i).ToArray()
		Next
		ScriptLogicalLinesArr = Lines
	End Property
	
	Property Get ScriptWords()
		Set ScriptWords = cScriptWords
	End Property
	
	Property Get ScriptWordsArr()
		ScriptWordsArr = cScriptWords.ToArray()
	End Property
	
	Property Get ScriptFullPath()
		ScriptFullPath = cScriptPath
	End Property
	
	Property Get ScriptCode()
		ScriptCode = Code
	End Property
	
	Public Function LoadWords(ByVal aWords)
		Set LoadWords = GetLogicalLines(aWords)
	End Function
	
	Public Function LoadCode(ByVal aCode)
		Dim WordSeperator: Set WordSeperator = New VBSPlus_WordSpliter
		Dim Words: Words = WordSeperator.Scan(aCode)
		Set LoadCode = LoadWords(Words)
	End Function
	
	Public Function LoadScript(ByVal aPath)
		' Deal Arguements
		If FSO.FolderExists(aPath)   Then Call PopupDebugInfo("LoadScript Error - Path is a Folder", "VBSPlus Interpreter", "Error", 0, 0)
		If Not FSO.FileExists(aPath) Then Call PopupDebugInfo("LoadScript Error - File Not Found"  , "VBSPlus Interpreter", "Error", 0, 0)
		
		' Read Script
		Dim AllCode: AllCode = FSO.OpenTextFile(aPath).ReadAll
		
		' Deal Script
		Set LoadScript = LoadCode(AllCode)
		cScriptPath = aPath
		cScriptShortPath = FSO.GetFile(aPath).ShortPath
	End Function
	
	Private cWords               ' Script Words Splited By VBSPlus_WordSpliter
	Private cScriptWords         ' Parsed Script Words
	Private ci                   ' Index
	Private cLineContinue        ' Continue This Line, Do Not Split New Line
	Private cThisLine            ' Current Line (Array), (In Future: (Object) {Words: [...], Type: \d})
	Private cLogicalLines        ' Result Lines
	
	Private Function InitVars(ByVal aWords)
		cWords        = aWords
		cLineContinue = False
		Set cScriptWords = New VbsList
		Set cThisLine = New VbsList
		Set cLogicalLines = New VbsList
	End Function
	
	' [X] Split Logical Lines
	' [X] Expand Compressed If
	' [X] Expand Compressed Case
	' [X] Clear Comments
	' [ ] Parse Line Type, Get Line Properties
	
	Public Function GetLogicalLines(ByVal aWords)
		Dim Word
		InitVars aWords
		
		For ci = 0 To UBound(cWords)
			Set Word = cWords(ci)
			Call NextWord(Word)
		Next
		Call NextLine()
		
		Set GetLogicalLines = cLogicalLines
	End Function
	
	Private Function NextWord(ByVal aWord)
		Dim WordType:  WordType  = aWord("Type")
		Dim WordValue: WordValue = aWord("Value")
		Dim UCValue:   UCValue   = UCase(WordValue)
		
		' Clear Useless Comments
		If ClearREM And WordType = 6 Then Exit Function
		
		' Line Continue Using "_"
		If WordValue = "_" Then cLineContinue = True: Exit Function
		
		If WordType = 0 Then
			' Physical Line Ends
			If Not cLineContinue Then
				' Logical Line Ends
				Call NextLine()
			End If
			
			cLineContinue = False
			Exit Function
		End If
		
		' Line Split Using ":"
		If WordValue = ":" Then Call NextLine(): Exit Function
		
		' Add This Word To This Line
		cThisLine.Push aWord
		cScriptWords.Push aWord
	End Function
	
	Private Function NextLine()
		If cThisLine.Length > 0 Then
			' Expand Compressed Statements
			Dim Lines ' <List> All New Lines
			Set Lines = ExpandCompressed(cThisLine)

			' Add New Lines To cLogicalLines
			Dim Line
			For Each Line In Lines.ToArray()
				cLogicalLines.Push(Line)
			Next
		End If
		
		' Prepare Next Line
		cThisLine.Clear
	End Function
	
	Private Function ExpandCompressed(ByVal aLine)
		Dim Line, Lines
		Dim i, ReduceWord
		Dim IfLayer, IfCompressed
		Set Line = New VbsList
		Set Lines = New VbsList
		IfCompressed = False
		
		For i = 0 To aLine.UBound
			ReduceWord = False
			
			' If
			If GetUValue(aLine, i) = "IF" Then IfLayer = IfLayer + 1
			
			' End If
			' Split Line Before AND After "End If" when if compressed
			If GetUValue(aLine, i) = "END" Then
				If GetUValue(aLine, i+1) = "IF" And IfCompressed Then
					' Split Line
					Lines.Push CloneList(Line)
					Line.Clear
					
					' Add "End" "If" To Line
					Line.Push aLine(i + 0)
					Line.Push aLine(i + 1)
					
					' Split Line
					Lines.Push CloneList(Line)
					Line.Clear
					
					' If Layer Decrease
					IfLayer = IfLayer - 1
					
					' Skip Following Word "IF"
					i = i + 1
					
					' No Repeat Word Appending
					ReduceWord = True
				End If
			End If
			
			' Then
			If GetUValue(aLine, i) = "THEN" Then
				' Add "Then" To Current Line
				Line.Push aLine(i)
				
				' Split Line
				Lines.Push CloneList(Line)
				Line.Clear
				
				' Word after "Then" means if compressed
				If aLine.UBound > i Then IfCompressed = True
				
				' No Repeat Word Appending
				ReduceWord = True
			End If
			
			' Else
			If GetUValue(aLine, i) = "ELSE" Then
				' Split Line
				Lines.Push CloneList(Line)
				Line.Clear
				
				' Add "Else" To Current Line
				Line.Push(aLine(i))
				
				' Split Line
				Lines.Push CloneList(Line)
				Line.Clear
				
				' No Repeat Word Appending
				ReduceWord = True
			End If
			
			' Case
			' Use "Do ... Exit Do ... Loop" As "If ... Exit If ... End If"
			Do While GetUValue(aLine, i) = "CASE"
				If i > 0 Then If GetUValue(aLine, i-1) = "SELECT" Then Exit Do
				
				' Add Word "case" to This Line
				Line.Push aLine(i)
				
				Const Connecters = "+ - * / \ ^ <> >= <= > < & , ( ) MOD IS NOT OR XOR AND EQV IMP"
				Dim Connecting
				Dim QuoteLayer: QuoteLayer = 0
				Dim j
				For j = i+1 To aLine.UBound
					' Add This Word
					Line.Push aLine(j)
					
					' Check Quote Layers
					If GetUValue(aLine, j) = "(" Then QuoteLayer = QuoteLayer + 1
					If GetUValue(aLine, j) = ")" Then QuoteLayer = QuoteLayer - 1
					
					' Check Connection
					Connecting = False
					If QuoteLayer = 0 And InStr(1, Connecters, GetUValue(aLine, j)) > 0 Then Connecting = True
					If aLine.UBound > j Then If QuoteLayer = 0 And InStr(1, Connecters, GetUValue(aLine, j+1)) > 0 Then Connecting = True
					
					' Split Line if no connection
					If Not Connecting And QuoteLayer = 0 Then
						' Split Line
						Lines.Push CloneList(Line)
						Line.Clear
						Exit For
					End If
				Next
				
				i = j
				ReduceWord = True
				Exit Do
			Loop
			
			If Not ReduceWord Then
				' Add This Word to Line
				Line.Push aLine(i)
			End If
		Next
		
		' Append Last Line
		If Line.Length > 0 Then
			Lines.Push CloneList(Line)
			Line.Clear
		End If
		
		' Fill Reduced "End If"
		If IfCompressed Then
			Dim WordEnd, WordIf
			For i = 1 To IfLayer
				Set WordEnd = CreateObject("Scripting.Dictionary"): WordEnd.Add "Type", 1.1: WordEnd.Add "Value", "End"
				Set WordIf  = CreateObject("Scripting.Dictionary"): WordIf. Add "Type", 1.1: WordIf. Add "Value", "If"
				Line.Clear
				Line.Push WordEnd
				Line.Push WordIf
				Lines.Push CloneList(Line)
				Line.Clear
			Next
		End If
		
		Set ExpandCompressed = Lines
	End Function
	
	' ## Support Functions ##
	
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
		If Source     = ""      Then Source = "VBSPlus LineParser Ver_" & LineParser_Version
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
	
	' Clone a VbsList
	Private Function CloneList(List)
		Set CloneList = (New VbsList).FromArray(List.ToArray())
	End Function
	
	' Get The UpperCased Value of Word from given Words and Index
	' Please make sure aWords(aIndex) Exists, This Function will NOT check it
	Private Function GetUValue(ByVal aWords, ByVal aIndex)
		GetUValue = UCase(aWords(aIndex)("Value"))
	End Function
	
	' Get Word that has not been scanned. Number + ci = Word's index.
	' If Not 0 <= Number + ci <= UBound(cWords) Then Returns {Type: -1, Value: ""}
	Private Function GetFutureWord(ByVal Number)
		Dim i: i = ci + Number
		Dim Word
		
		If i > UBound(cWords) Or i < 0 Then
			Set Word = CreateObject("Scripting.Dictionary")
			Word.Add "Type", -1
			Word.Add "Value", ""
		Else
			Set Word = cWords(i)
		End If
		
		Set GetFutureWord = Word
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
		For i = 0 To UBound(Arr)
			If Arr(i) = Item Then
				ItemInArray = True
				Exit Function
			End If
		Next
	End Function
End Class









Class VBSPlus_Interpreter
	Private Sub Class_Initailize()
		Set cFSO = CreateObject("Scripting.FileSystemObject")
	End Sub
	
	Private Sub Class_Terminate()
		'
	End Sub
	
	Private cFSO
	
	Private Function LoadScript(ByVal aPath)
		If Not cFSO.FileExists Then LoadScript = -1: Exit Function
		Dim LP: Set LP = New VBSPlus_LineParser
		Dim Lines: Lines = LP.LoadScript(aPath)
		LoadWords = LoadLines(Lines)
	End Function
	
	Private Function LoadCode(ByVal aCode)
		Dim LP: Set LP = New VBSPlus_LineParser
		Dim Lines: Lines = LP.LoadCode(aCode)
		LoadWords = LoadLines(Lines)
	End Function
	
	Private Function LoadWords(ByVal aWords)
		Dim LP: Set LP = New VBSPlus_LineParser
		Dim Lines: Lines = LP.LoadWords(aWords)
		LoadWords = LoadLines(Lines)
	End Function
	
	Private Function LoadLines(ByVal aLines)
		LoadLines = ExecuteVBS(aLines)
	End Function
	
	Private cLines
	
	Private Function InitVars(ByVal aLines)
		cLines = aLines
	End Function
	
	Private Function ExecuteVBS(ByVal aLines)
		InitVars aLines
		GetFunctions cLines
	End Function
	
	' Get Functions Like [{Name: "", Arguments: ["",], Statments: ["",]},]
	Private Function GetFunctions(ByVal aLines)
		Dim i, Line, Values
		Dim ThisFunc, Functions(), FUNC_UBound
		Dim FuncStart
		FUNC_UBound = 0: ReDim Functions(0)
		InFunc = False
		
		For i = 0 To UBound(aLines)
			Line = aLines(i)
			Values = GetLineValues(Line, True)
			
			' Function Start
			If Values(0) = "FUNCTION" Then
				Set ThisFunc = CreateObject("Scripting.Dictionary")
				ThisFunc.Add "Name", Values(2)
				ThisFunc.Add "Arguments", GetFuncArgs(Line)
			End If
			
			If (Values(0) = "PUBLIC" Or Values(0) = "PRIVATE") And Values(1) = "FUNCTION" Then
				Set ThisFunc = CreateObject("Scripting.Dictionary")
				ThisFunc.Add "Name", Values(3)
				ThisFunc.Add "Arguments", GetFuncArgs(Line)
			End If
			
			' Function End
			If Values(0) = "End" And Values(1) = "FUNCTION" Then
				ReDim Preserve Functions(FUNC_UBound)
				ThisFunc.Add "Statments", SliceArray(aLines, FuncStart, i)
				Set Functions(FUNC_UBound) = ThisFunc
				FUNC_UBound = FUNC_UBound + 1
			End If
		Next
	End Function
	
	Private Function GetFuncArgs(ByVal aLine)
		Dim i, Word, Values, InArgs
		Dim wType, wValue
		Dim Args(), Args_UBound
		Values = GetLineValues(aLine, True)
		InArgs = False
		ReDim Args(0): Args_UBound = 0
		
		For i = 0 To UBound(aLine)
			Set Word = aLine(i)
			wType  = Word("Type")
			wValue = Word("Value")
			If Value = "(" Then InArgs = True
			If Value = ")" Then InArgs = False: Exit For
			If InArgs And wType = 1.0 Then
				ReDim Preserve Args(Args_UBound)
				Args(Args_UBound) = wValue
				Args_UBound = Args_UBound + 1
			End If
		Next
		
		GetFuncArgs = Args
	End Function
	
	' ## Support Functions ##
	' Get All Word Values of a Line
	Private Function GetLineValues(ByVal aLine, ByVal UpperCase)
		Dim i, Word, Values()
		ReDim Values(UBound(aLine))
		For i = 0 To UBound(aLine)
			Set Word = aLine(i)
			Values(i) = Word("Value")
			If UpperCase Then Values(i) = UCase(Values(i))
		Next
		GetLineValues = Values
	End Function
	
	' Get part of an existing array
	' Contains Arr(FromIndex) But Not Arr(ToIndex)
	Private Function SliceArray(ByVal Arr, ByVal FromIndex, ByVal ToIndex)
		Dim i, j, NewArr(): ReDim NewArr(ToIndex-FromIndex-1): j = 0
		
		For i = FromIndex To ToIndex-1
			NewArr(j) = Arr(i)
			j = j + 1
		Next
		
		SliceArray = NewArr
	End Function
End Class









Class VbsList
	Private Sub Class_Initialize()
		ListUBound = -1
	End Sub
	
	Private ListArray()
	Private ListUBound
	
	Public Function Push(ByVal Item)
		ListUBound = ListUBound + 1
		ReDim Preserve ListArray(ListUBound)
		If IsObject(Item) Then
			Set ListArray(ListUBound) = Item
		Else
			ListArray(ListUBound) = Item
		End If
	End Function
	
	Public Default Function GetItem(ByVal Index)
		If IsObject(ListArray(Index)) Then
			Set GetItem = ListArray(Index)
		Else
			GetItem = ListArray(Index)
		End If
	End Function
	
	Public Function SetItem(ByVal Index, ByVal Value)
		ListArray(Index) = Value
	End Function
	
	Public Function Clear()
		Erase ListArray
		ListUBound = -1
	End Function
	
	Public Function ToArray()
		ToArray = ListArray
	End Function
	
	Public Function FromArray(ByRef Arr)
		Dim Item
		For Each Item In Arr
			Push(Item)
		Next
		Set FromArray = Me
	End Function
	
	Public Property Get Length()
		Length = ListUBound + 1
	End Property
	
	Public Property Get UBound()
		UBound = ListUBound
	End Property
End Class

Class VbsStack
	Private StackArray(), StackUBound
	
	Private Sub Class_Initialize()
		StackUBound = -1
	End Sub
	
	Public Function Push(ByVal Item)
		StackUBound = StackUBound + 1
		ReDim Preserve StackArray(StackUBound)
		If IsObject(Item) Then
			Set StackArray(StackUBound) = Item
		Else
			StackArray(StackUBound) = Item
		End If
	End Function
	
	Public Function Pop()
		If StackUBound = -1 Then Err.Raise 9001, "VbsStack", "Trying to Pop While Nothing In Stack"
		Pop = StackArray(StackUBound)
		StackUBound = StackUBound - 1
		ReDim Preserve StackArray(StackUBound)
	End Function
	
	Public Function Clear()
		StackUBound = -1
		Erase StackArray
	End Function
	
	Public Function ToArray()
		ToArray = StackArray(StackUBound)
	End Function
	
	Property Get LastItem()
		LastItem = StackArray(StackUBound)
	End Property
	
	Property Get Count()
		Count = StackUBound + 1
	End Property
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