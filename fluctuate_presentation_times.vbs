'�I�����ꂽ�X���C�h�̃v���[�����Ԃ��u���₷�E���炷�v�}�N��
'
'�i����͂ȂɁj
'�u���n�[�T���@�\�v�Ŋe�X���C�h�Ƀv���[�����Ԃ�ݒ肵���Ƃ��ɁA
'�O���̎��Ԃ����炵�āA���̕��A�㔼�֊���U��Ƃ����p�r�Ŏg���܂��B
'
'�X���C�h�̖����������Ƃ��Ɂu���n�[�T���@�\�v�����x���g���͔̂�����Ȃ̂Ŗ{�}�N�������܂����B
'
'
'�i�g�����j
'�E�܂����߂ɁA�v���[�����Ԃ𑝌��������X���C�h��I�����܂��B
'�E�}�N�����N������ƃe�L�X�g����͂���a�n�w������܂��A
'�E5 �Ɠ��͂��n�j�{�^���������ƁA�v���[���̎��Ԃ�5�b���Z����܂��B
'  �i���̎��A-5 ����͂��Ă���ƁA5�b���Z���܂��B�j
'
'

Option Explicit

Sub FluctuatePresentationTimes()
	Dim Value As Integer
	Dim Text As String
	Dim Slide As Slide
	Dim Time As Double

	Text = InputBox("�������������ԁi�b�j����͂��ĉ������B")
	If False = IsNumeric(Text) Then
		MsgBox "���l�ł͂Ȃ����������͂���܂����A���l����͂��ĉ������B"
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
