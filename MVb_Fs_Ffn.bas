Attribute VB_Name = "MVb_Fs_Ffn"
Option Compare Binary
Const CMod$ = "MVb_Fs_Ffn."

Option Explicit

Function CutExt$(A)
CutExt = Ffn_XCut_Ext(A)
End Function

Function CvFfnAy(FfnAy0) As String()
Select Case True
Case IsStr(FfnAy0):   CvFfnAy = Ap_Sy(FfnAy0)
Case IsArray(FfnAy0): CvFfnAy = Ay_Sy(FfnAy0)
Case Else: Stop
End Select
End Function

Function FfnAy_Chk_NotExist(FfnAy0) As String()
Dim I, O$()
For Each I In CvFfnAy(FfnAy0)
    If Not Ffn_Exist(CStr(I)) Then
        PushI O, I
    End If
Next
If Sz(O) = 0 Then Exit Function
FfnAy_Chk_NotExist = MsgAp_Ly("[File(s)] not found", O)
End Function

Sub FfnAy_NotExist_XAss(FfnAy0, Fun$)
Dim I, O$()
For Each I In CvFfnAy(FfnAy0)
    If Ffn_Exist(CStr(I)) Then
        PushI O, I
    End If
Next
If Sz(O) = 0 Then Exit Sub
XThw Fun, "File(s) should not exist", "Files", O
End Sub

Sub FfnAy_XAss_Exist(FfnAy$(), Fun$)
Dim O$(), I
For Each I In FfnAy
    If Not Fso.FileExists(I) Then PushI O, I
Next
If Sz(O) = 0 Then Exit Sub
XThw Fun, "File(s) not found", "File(s)", O
End Sub

Function FfnAy_XCpy_ToPth_IfDif(FfnAy0, Pth$) As String()
Pth_XEns Pth
Dim I, O$()
For Each I In CvFfnAy(FfnAy0)
    PushNonBlankStr O, Ffn_XCpy_ToPthIfDif(I, Pth)
Next
If Sz(O) > 0 Then
    PushMsgUnderLinDbl O, "Above files not found"
    FfnAy_XCpy_ToPth_IfDif = O
End If
End Function

Sub FfnAy_XDlt_IfExist(A)
Ay_XDo A, "Ffn_XDltIfExist"
End Sub

Function FfnAy_XWh_Exist(A) As String()
Dim Ffn
For Each Ffn In AyNz(A)
    If Ffn_Exist(Ffn) Then
        PushI FfnAy_XWh_Exist, Ffn
    End If
Next
End Function

Function FfnDTim$(A)
If Ffn_Exist(A) Then
    FfnDTim = Dte_DTim(FileDateTime(A))
End If
End Function

Function FfnStampDr(A) As Variant()
FfnStampDr = Array(A, Ffn_Sz(A), Ffn_Tim(A), Now)
End Function

Function FfnXEnsPth$(A)
Pth_XEnsAll Ffn_Pth(A)
FfnXEnsPth = A
End Function

Function Ffn_Atr(A, OAtr&) As Boolean
OAtr = GetAttr(A)
End Function

Function Ffn_Chk_NotExist(A$) As String()
If Ffn_Exist(A) Then Exit Function
Ffn_Chk_NotExist = MsgAp_Ly("[File] not exist", A)
End Function

Function Ffn_Exist(A) As Boolean
Ffn_Exist = Fso.FileExists(A)
End Function

Function Ffn_Ext$(A)
Dim B$, P%
B = Ffn_Fn(A)
P = InStrRev(B, ".")
If P = 0 Then Exit Function
Ffn_Ext = Mid(B, P)
End Function

Function Ffn_Fdr$(A)
Ffn_Fdr = Pth_Fdr(Ffn_Pth(A))
End Function

Function Ffn_Fn$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then Ffn_Fn = A: Exit Function
Ffn_Fn = Mid(A, P + 1)
End Function

Function Ffn_Fnn$(A)
Ffn_Fnn = Ffn_XCut_Ext(Ffn_Fn(A))
End Function

Function Ffn_IsSam(A, B) As Boolean
If Ffn_Tim(A) <> Ffn_Tim(B) Then Exit Function
If Ffn_Sz(A) <> Ffn_Sz(B) Then Exit Function
Ffn_IsSam = True
End Function

Function Ffn_MsgLy_AlreadyLoaded(A$, FilKind$, LdTim$) As String()
Dim Sz&, Tim$, Ld$, Msg$
Sz = Ffn_Sz(A)
Tim = FfnDTim(A)
Msg = QQ_Fmt("[?] file of [time] and [size] is already loaded [at].", FilKind)
Ffn_MsgLy_AlreadyLoaded = MsgAp_Ly(Msg, A, Tim, Sz, LdTim)
End Function

Function Ffn_Msg_IsSam(A, B, Sz&, Tim$, Optional Msg$) As String()
Dim O$()
Push O, "File 1   : " & A
Push O, "File 2   : " & B
Push O, "File Size: " & Sz
Push O, "File Time: " & Tim
Push O, "File 1 and 2 have same size and time"
If Msg <> "" Then Push O, Msg
Ffn_Msg_IsSam = O
End Function

Function Ffn_Msg_NotExist(A, Optional FilKind$ = "File") As String()
Dim F$
F = QQ_Fmt("[?] not found in [folder]", FilKind)
Ffn_Msg_NotExist = MsgAp_Ly(F, Ffn_Fn(A), Ffn_Pth(A))
End Function

Function Ffn_NotExist(A) As Boolean
Ffn_NotExist = Not Ffn_Exist(A)
End Function

Sub Ffn_NotExistAss(Ffn, Fun$, Optional FilKd$)
If Ffn_Exist(Ffn) Then XThw Fun, "File should not exist", "File-Pth File-Name File-Kind", Ffn_Pth(Ffn), Ffn_Fn(Ffn), FilKd
End Sub

Function Ffn_NxtFfn$(A$)
If Not Ffn_Exist(A) Then Ffn_NxtFfn = A: Exit Function
Dim J%, O$
For J = 1 To 999
    O = Ffn_XAdd_FnSfx(A, "(" & Format(J, "000") & ")")
    If Not Ffn_Exist(O) Then Ffn_NxtFfn = O: Exit Function
Next
Stop
End Function

Function Ffn_Pth$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then Exit Function
Ffn_Pth = Left(A, P)
End Function

Function Ffn_Sz&(A)
If Not Ffn_Exist(A) Then Ffn_Sz = -1: Exit Function
Ffn_Sz = FileLen(A)
End Function

Function Ffn_Tim(A) As Date
If Ffn_Exist(A) Then Ffn_Tim = FileDateTime(A)
End Function

Function Ffn_TimSzStr$(A)
If Ffn_Exist(A) Then Ffn_TimSzStr = FfnDTim(A) & "." & Ffn_Sz(A)
End Function

Function Ffn_XAdd_FnPfx$(A$, Pfx$)
Ffn_XAdd_FnPfx = Ffn_Pth(A) & Pfx & Ffn_Fn(A)
End Function

Function Ffn_XAdd_FnSfx$(A$, Sfx$)
Ffn_XAdd_FnSfx = Ffn_XRmv_Ext(A) & Sfx & Ffn_Ext(A)
End Function

Sub Ffn_XAsg_Tim_Sz(A$, OTim As Date, OSz&)
If Not Ffn_Exist(A) Then
    OTim = 0
    OSz = 0
    Exit Sub
End If
OTim = Ffn_Tim(A)
OSz = Ffn_Sz(A)
End Sub

Sub Ffn_XAss_Exist(A, Fun$, Optional FilKd$ = "File")
If Not Fso.FileExists(A) Then
    Dim M$: M = QQ_Fmt("? not found", FilKd)
    XThw Fun, M, "File File-Path File-Name CurPath", A, Ffn_Pth(A), Ffn_Fn(A), CurPth
End If
End Sub

Sub Ffn_XAss_NotExist(A, Fun$)
If Not Fso.FileExists(A) Then XThw Fun, "File should not exist", "Path FileName CurPth Sz Tim", Ffn_Pth(A), Ffn_Fn(A), CurPth, Ffn_Sz(A), Ffn_Tim(A)
End Sub

Sub Ffn_XCpy(A, ToFfn$, Optional OvrWrt As Boolean)
If OvrWrt Then Ffn_XDlt ToFfn
Fso.GetFile(A).Copy ToFfn
End Sub

Function Ffn_XCpy_ToNxtFfn$(A$)
Dim O$
O = Ffn_NxtFfn(A)
Ffn_XCpy A, O
Ffn_XCpy_ToNxtFfn = O
End Function

Sub Ffn_XCpy_ToPth(A, ToPth$, Optional OvrWrt As Boolean)
Fso.CopyFile A, ToPth$ & Ffn_Fn(A), OvrWrt
End Sub

Function Ffn_XCpy_ToPthIfDif(A, Pth$) As Boolean
Const M_Sam$ = "File is same the one in Path."
Const M_Copied$ = "File is copied to Path."
Const M_NotFnd$ = "File not found, cannot copy to Path."
Dim B$, Msg$
Select Case True
Case Ffn_Exist(A)
    B = Pth & Ffn_Fn(A)
    Select Case True
    Case Ffn_IsSam(B, A)
        Msg = M_Sam: GoSub Prt
    Case Else
        Fso.CopyFile A, B, True
        Msg = M_Copied: GoSub Prt
    End Select
Case Else
    Msg = M_NotFnd: GoSub Prt
    Ffn_XCpy_ToPthIfDif = A
End Select
Exit Function
Prt:
    Debug.Print QQ_Fmt("Ffn_XCpy_ToPthIfDif: ? Path=[?] File=[?]", Msg, Pth, A)
    Return
End Function

Function Ffn_XCut_Ext$(A)
Dim B$, C$, P%
B = Ffn_Fn(A)
P = InStrRev(B, ".")
If P = 0 Then
    C = B
Else
    C = Left(B, P - 1)
End If
Ffn_XCut_Ext = Ffn_Pth(A) & C
End Function

Sub Ffn_XDlt(A)
Const CSub$ = CMod & "Ffn_XDlt"
On Error GoTo X
Kill A
Exit Sub
X:
Dim E$
E = Err.Description
XThw CSub, "Cannot kill", "Ffn Er", A, E
End Sub

Sub Ffn_XDltIfExist(A)
If Ffn_Exist(A) Then Ffn_XDlt A
End Sub

Sub Ffn_XMov_ToPth(A$, ToPth$)
Fso.MoveFile A, ToPth
End Sub

Function Ffn_XRmv_Ext$(A)
Ffn_XRmv_Ext = Ffn_XCut_Ext(A)
End Function

Function Ffn_XRpl_Ext$(A, NewExt)
Ffn_XRpl_Ext = Ffn_XRmv_Ext(A) & NewExt
End Function

Private Sub ZZ()
Dim A
Dim B$
Dim C As Date
Dim D&
Dim E As Boolean
Dim F$()
Dim XX
CutExt A
CvFfnAy A
FfnAy_Chk_NotExist A
FfnAy_NotExist_XAss A, B
FfnAy_XAss_Exist F, B
FfnAy_XCpy_ToPth_IfDif A, B
FfnAy_XDlt_IfExist A
FfnAy_XWh_Exist A
FfnDTim A
FfnStampDr A
FfnXEnsPth A
Ffn_Atr A, D
Ffn_Chk_NotExist B
Ffn_Exist A
Ffn_Ext A
Ffn_Fdr A
Ffn_Fn A
Ffn_Fnn A
Ffn_IsSam A, A
Ffn_MsgLy_AlreadyLoaded B, B, B
Ffn_Msg_IsSam A, A, D, B, B
Ffn_NotExist A
Ffn_NotExistAss A, B, B
Ffn_NxtFfn B
Ffn_Pth A
Ffn_Sz A
Ffn_Tim A
Ffn_TimSzStr A
Ffn_XAdd_FnPfx B, B
Ffn_XAdd_FnSfx B, B
Ffn_XAsg_Tim_Sz B, C, D
Ffn_XAss_Exist A, B
Ffn_XAss_NotExist A, B
Ffn_XCpy A, B, E
Ffn_XCpy_ToNxtFfn B
Ffn_XCpy_ToPth A, B, E
Ffn_XCpy_ToPthIfDif A, B
Ffn_XCut_Ext A
Ffn_XDlt A
Ffn_XDltIfExist A
Ffn_XMov_ToPth B, B
Ffn_XRmv_Ext A
Ffn_XRpl_Ext A, A
End Sub

Private Sub Z()
End Sub
