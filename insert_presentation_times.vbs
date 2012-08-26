'プレゼンの進捗時間をメモへ挿入する
'
'（マクロの目的）
'与えられた時間ちょうどでプレゼンを終わらせる。
'そのために、ノートに進捗時間を書き込む
'
'（時間の書式）
'【time】00:00 → 00:09
'
'（使い方）
'まず、パワーポイントの「リハーサル機能」を使い各スライドの所用時間を記録します。
'次に、本マクロを実行します。
'メモに時間が書き込まれているので確認して下さい。
'
'

Option Explicit

Sub InsertPresentationTimes()

	Dim objSld As Slide
	Dim lngStartTime As Long
	Dim lngFinishTime As Long
	Dim memotext As String
	Dim StartText As String
	Dim FinishText As String

	lngStartTime = 0
	lngFinishTime = 0
	For Each objSld In ActivePresentation.Slides
		lngFinishTime = lngFinishTime + objSld.SlideShowTransition.AdvanceTime
		memotext = RemoveTimeLine(objSld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange)
		StartText = FormatTime(lngStartTime)
		FinishText = FormatTime(lngFinishTime)
		memotext = GetTagName() & StartText & " → " & FinishText & vbCrLf & memotext

		objSld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange = memotext
		lngStartTime = lngStartTime + objSld.SlideShowTransition.AdvanceTime
	Next objSld
End Sub

Function FormatTime(ByVal sec As Long) As String
	Dim s As String
	Dim m As String
	s = SupplementZero(sec Mod 60)
	m = SupplementZero(sec \ 60)
	FormatTime = m & ":" & s
End Function

Function SupplementZero(ByVal Text As Long) As String
	Dim n As Long
	n = Len(CStr(Text))
	If n = 0 Then
		SupplementZero = "00"
	ElseIf n = 1 Then
		SupplementZero = "0" & Text
	Else
		SupplementZero = Text
	End If
End Function

Function RemoveTimeLine(ByVal Text As String) As String
	Dim pos As Long
	Dim newtext As String
	Dim retpos As Long

	pos = InStr(1, Text, GetTagName())
	If pos = 1 Then
		'先頭にタグが挿入されている
		retpos = InStr(Text, Chr(13))
		If retpos <> 0 Then
			newtext = MyLtrim(Mid(Text, retpos))
		Else
			newtext = ""
		End If
	Else
		newtext = Text
	End If
	RemoveTimeLine = newtext
End Function

Function MyLtrim(ByVal Text As String) As String
	Dim c As String
	Dim newtext As String

	c = Mid(Text, 1, 1)
	If (c = Chr(13)) Or (c = Chr(10)) Then
		newtext = Mid(Text, 2)
	Else
		newtext = Text
	End If
	MyLtrim = newtext
End Function


Function GetTagName() As String
	Dim x As String
	x = "【time】"
	GetTagName = x
End Function


