Attribute VB_Name = "MVb_Fs_Pth"
Option Compare Binary
Option Explicit
Public Const PthSep$ = "\"

Function Cd$(Optional A$)
If A = "" Then
    Cd = Pth_XEns_Sfx(CurDir)
    Exit Function
End If
ChDir A
Cd = Pth_XEns_Sfx(A)
End Function

Function CvPth$(ByVal A$)
If A = "" Then
    A = CurDir
End If
CvPth = Pth_XEns_Sfx(A)
End Function

Property Get CurPth$()
CurPth = Cd
End Property
Sub Pth_XBrw(A$)
Shell QQ_Fmt("Explorer ""?""", A), vbMaximizedFocus
End Sub

Sub Pth_XClr(A$)
FfnAy_XDlt_IfExist Pth_FfnAy(A)
End Sub

Sub Pth_XClr_Fil(A$)
If Not PthIsExist(A) Then Exit Sub
Dim F
For Each F In AyNz(Pth_FfnAy(A))
   Ffn_XDlt F
Next
End Sub
Function PthXHas_Sfx(A) As Boolean
PthXHas_Sfx = XTak_LasChr(A) = PthSep
End Function
Function Pth_XEns_Sfx$(A)
If PthXHas_Sfx(A) Then
    Pth_XEns_Sfx = A
Else
    Pth_XEns_Sfx = A & PthSep
End If
End Function
Function Pth_XEns$(A$)
If Not Fso.FolderExists(A) Then MkDir A
Pth_XEns = Pth_XEns_Sfx(A)
End Function

Function Pth_XEnsAll$(A$)
Dim Ay$(): Ay = Split(A, PthSep)
Dim J%, O$
O = Ay(0)
For J = 1 To UB(Ay)
    O = O & PthSep & Ay(J)
    Pth_XEns O
Next
Pth_XEnsAll = A
End Function

Function Pth_EntAy(A$, Optional FilSpec$ = "*.*", Optional Atr As FileAttribute) As String()
Stop
'Erase O
End Function

Function Pth_Fdr$(A$)
Pth_Fdr = XTak_AftRev(XRmv_LasChr(A), "\")
End Function

Function Pth_FfnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Pth_FfnAy = Ay_XAdd_Pfx(Pth_FnAy(A, Spec, Atr), A)
End Function

Function Pth_FnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Ass PthIsExist(A)
Ass PthHasPthSfx(A)
If Atr And vbDirectory Then Stop
Dim O$()
Dim M$
M = Dir(A & Spec, Atr)
While M <> ""
   PushI Pth_FnAy, M
   M = Dir
Wend
End Function

Function Pth_FxAy(A$) As String()
Dim O$(), B$
If Right(A, 1) <> "\" Then Stop
B = Dir(A & "*.xls")
Dim J%
While B <> ""
    J = J + 1
    If J > 1000 Then Stop
    If Ffn_Ext(B) = ".xls" Then
        Push O, A & B
    End If
    B = Dir
Wend
Pth_FxAy = O
End Function

Function PthHasFil(A) As Boolean
PthHasFil = Fso.GetFolder(A).Files.Count > 0
End Function

Function PthHasPthSfx(A) As Boolean
PthHasPthSfx = XTak_LasChr(A) = "\"
End Function

Function PthHasFdr(A) As Boolean
PthHasFdr = Fso.GetFolder(A).SubFolders.Count > 0
End Function

Function PthIsEmp(A) As Boolean
If PthHasFil(A) Then Exit Function
If PthHasFdr(A) Then Exit Function
PthIsEmp = True
End Function

Function PthIsExist(A) As Boolean
PthIsExist = Fso.FolderExists(A)
End Function

Sub PthMovFilUp(A$)
Dim I, Tar$
Tar$ = PthUp(A)
For Each I In AyNz(Pth_FnAy(A))
    Ffn_XMov_ToPth CStr(I), Tar
Next
End Sub

Sub PthRenXAdd_Pfx(A, Pfx)
PthRen A, PthXAdd_Pfx(A, Pfx)
End Sub
Function PthXAdd_Pfx(A, Pfx)
With Brk2Rev(RmvSfx(A, PthSep), PthSep, NoTrim:=True)
    PthXAdd_Pfx = .S1 & PthSep & Pfx & .S2 & PthSep
End With
End Function
Sub PthRen(A, NewPth)
If PthIsExist(NewPth) Then XThw CSub, "NewPth exist", "Pth NewPth", A, NewPth
If Not PthIsExist(A) Then XThw CSub, "Pth not exist", "Pth NewPth", A, NewPth
Fso.GetFolder(A).Name = NewPth
End Sub

Function Pth_FdrAy(A, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
'Pth_FdrAy = Itr_Ny(Fso.GetFolder(A).SubFolders, Spec)
'Exit Function
Ass PthIsExist(A)
Ass PthHasPthSfx(A)
Dim M$, X&, Atr1&
X = Atr Or vbDirectory
M = Dir(A & Spec, vbDirectory)
While M <> ""
    If InStr(M, "?") > 0 Then
        XDmp_Ly CSub, "Unicode entry is skipped", "UniCode-Entry Pth Spec Atr", M, A, Spec, Atr
        GoTo Nxt
    End If
    If M = "." Then GoTo Nxt
    If M = ".." Then GoTo Nxt
    If Ffn_Atr(A & M, Atr1) Then
        If Atr1 And X Then
            PushI Pth_FdrAy, M
        End If
    End If
Nxt:
    M = Dir
Wend
End Function

Private Sub Z_Pth_PthAy()
Dim Pth$
Pth = "C:\Users\user\AppData\Local\Temp\Adobe_ADMLogs\"
Ept = Ap_Sy()
GoSub Tst
Exit Sub
Tst:
    Act = Pth_PthAy(Pth)
    C
    Return
End Sub
Function Pth_PthAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Pth_PthAy = Ay_XAdd_PfxSfx(Pth_FdrAy(A, Spec, Atr), A, "\")
End Function

Function PthUp$(A, Optional Up% = 1)
Dim O$, J%
O = A
For J = 1 To Up
    O = PthUpOne(O)
Next
PthUp = O
End Function

Function PthUpOne$(A$)
PthUpOne = XTak_BefOrAllRev(RmvSfx(A, "\"), "\") & "\"
End Function

Private Sub ZZ_Pth_FxAy()
Dim A$()
A = Pth_FxAy(CurDir)
Ay_XDmp A
End Sub

Private Sub ZZ_Pth_XRmv_EmpSubDir()
Pth_XRmv_EmpSubDir TmpPth
End Sub

