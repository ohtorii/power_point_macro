'選択されたスライドのプレゼン時間を「増やす・減らす」マクロ
'
'（これはなに）
'「リハーサル機能」で各スライドにプレゼン時間を設定したときに、
'前半の時間を減らして、その分、後半へ割り振るという用途で使います。
'
'スライドの枚数が多いときに「リハーサル機能」を何度も使うのは非効率なので本マクロを作りました。
'
'
'（使い方）
'・まず初めに、プレゼン時間を増減したいスライドを選択します。
'・マクロを起動するとテキストを入力するＢＯＸが現れます、
'・5 と入力しＯＫボタンを押すと、プレゼンの時間が5秒加算されます。
'  （この時、-5 を入力していると、5秒減算します。）
'
'

Option Explicit

Sub FluctuatePresentationTimes()
	Dim Value As Integer
	Dim Text As String
	Dim Slide As Slide
	Dim Time As Double

	Text = InputBox("増減したい時間（秒）を入力して下さい。")
	If False = IsNumeric(Text) Then
		MsgBox "数値ではない文字が入力されました、数値を入力して下さい。"
	Else
		Value = CInt(Text)
		If 0 <> Value Then
			For Each Slide In Windows(1).Selection.SlideRange
				Time = Slide.SlideShowTransition.AdvanceTime
				Time = Time + Value
				If Time < 0 Then
					Time = 0
				End If
				Slide.SlideShowTransition.AdvanceTime = Time
			Next
		End If
	End If
End Sub
