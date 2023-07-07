VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_人臉蒐集報表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub start_time_AfterUpdate()
    end_time=start_time+6
End Sub
    
    Private Sub end_time_AfterUpdate()
    End Sub

Private  Sub time_CK(CK_mode As Boolean)'True:check start time False:check end time
    If (start_time > Date()) Then
        MsgBox "開始時間不可大於現在時間"
        If (CK_mode) Then 
            start_time = Date()
        Else
            end_time = Date()
        End If
    ElseIf (start_time > end_time) Then
        MsgBox "開始時間不可大於結束時間"
        If (CK_mode) Then 
            start_time = end_time
        Else
            end_time = start_time
        End If
    End If
End Sub
