'VBA code to find all common multiples of 2 numbers in Excel 2013
Sub directref_part1() 
Dim BaseCell As Range 
Set BaseCell = Range("b1") 
n = 0 
Do Until BaseCell.Offset(n, 0) = "" 
BaseCell.Offset(n, 0).Value2 = "" 
n = n + 1 
Loop 
End Sub
Sub direct_ref_part2() 
directref_part1 
Dim BaseCell As Range 
Dim BaseCell1 As Range 
Dim BaseCell2 As Range 
Set BaseCell = Range("b1") 
Set BaseCell1 = Range("a1") 
Set BaseCell2 = Range("a2") 
q = 2 
n = 0 
a1 = BaseCell1.Value2 
a2 = BaseCell2.Value2 
minval = a1 
If minval > a2 Then a2 = minval 
Do Until q = minval 
v1 = a1 / q 
v2 = a2 / q 
If v1 = Int(v1) Then 
If v2 = Int(v2) Then 
BaseCell.Offset(n, 0).Value2 = q 
n = n + 1 
End If 
End If 
q = q + 1 
Loop 
End Sub 
mouseevents 
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub w()
gamew
End Sub
Sub gamem()
For c = 1 To 8
For b = 1 To 8
SetCursorPos 320 + (b * 50), 300 + (c * 30)
For a = 1 To 10
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
Next a
Next b
Next c
End Sub
Sub gamew()
SetCursorPos 550, 450
For a = 1 To 40
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
Next a
End Sub
Sub clicker()
SetCursorPos 230, 750
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
Sleep 500
For x = 1 To 2000
SetCursorPos 600, 600
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
Sleep 2
Next x
End Sub
