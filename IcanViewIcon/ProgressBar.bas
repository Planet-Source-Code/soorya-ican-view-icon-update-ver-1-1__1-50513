Attribute VB_Name = "ProgressBar"
Option Explicit
'it is my funny and simple prog bar
'how to use   pbr me,Max,value


Public Sub pbr(ByRef WhichForm As Form, ByVal Max As Integer, ByVal Value As Integer)
    If Value + 30 <= Max Then
        WhichForm.txtHider.Width = WhichForm.imgProgressBar.Width - (((WhichForm.imgProgressBar.Width / Max) * Value)) - 30
        WhichForm.txtHider.Left = WhichForm.imgProgressBar.Left + ((WhichForm.imgProgressBar.Width / Max) * Value)
    End If
    
    If Value + 30 >= Max Then
         InitPB WhichForm
    Else
        WhichForm.txtCounter.Text = Round(CLng(Value) * 100 / Max) & "%"
    End If
End Sub


Public Sub InitPB(ByRef WhichForm As Form)
    WhichForm.txtHider.Left = WhichForm.imgProgressBar.Left + 10
    WhichForm.txtHider.Top = WhichForm.imgProgressBar.Top + 30
    WhichForm.txtHider.Height = WhichForm.imgProgressBar.Height - 60 '40
    WhichForm.txtHider.Width = WhichForm.imgProgressBar.Width - 40
    
    WhichForm.txtCounter.Left = (WhichForm.imgProgressBar.Width / 2) '+ (WhichForm.txtCounter.Width / 2)
    WhichForm.txtCounter.Top = (WhichForm.imgProgressBar.Top + WhichForm.imgProgressBar.Height / 2) - (WhichForm.txtCounter.Height / 2)
    
    WhichForm.txtCounter.Text = ""
End Sub

