Attribute VB_Name = "MVb__Tim"
Option Compare Binary
Option Explicit
Private M$, Beg As Date
Sub TimBeg(Optional Msg$ = "Time")
If M <> "" Then TimEnd
M = Msg
Beg = Now
End Sub
Sub TimEnd(Optional XHalt As Boolean)
Debug.Print M & " " & DateDiff("S", Beg, Now) & "(s)"
If XHalt Then Stop
End Sub
