Attribute VB_Name = "Module1"
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub SlowDown(MilliSeconds As Long)

Dim lngTickStore As Long

lngTickStore = GetTickCount()

Do While lngTickStore + MilliSeconds > GetTickCount()
DoEvents
Loop

End Sub

Public Function Pause(NumberOfSeconds As Variant)
On Error GoTo Err_Pause

    Dim PauseTime As Variant, start As Variant

    PauseTime = NumberOfSeconds
    start = Timer
    Do While Timer < start + PauseTime
    DoEvents
    Loop

Exit_Pause:
    Exit Function

Err_Pause:
    Resume Exit_Pause

End Function
