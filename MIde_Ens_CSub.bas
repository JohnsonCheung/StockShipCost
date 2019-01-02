Attribute VB_Name = "MIde_Ens_CSub"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_XEns_CSub."

Sub XEns_MdCSub()
WMd_XEns CurMd
End Sub

Sub XEns_PjCSub()
WPj_XEns CurPj
End Sub

Private Sub WMd_XEns(A As CodeModule)
With MdCSubBrk(A)
    WMd_XEns1 A, .MthBrkAy  '<== MthBrk must first
    WWMd_XEns2 A, .MdBrk
End With
End Sub

Private Sub WMd_XEns1(A As CodeModule, B() As CSubBrkMth) _
'sfdf _
'lksjdf
'lksdjf
Const CSub$ = CMod & "WMd_XEns1"
Const Trace As Boolean = True
Dim J%
WMd_XEns1a Md_Nm(A), B ' Ass if in sorting order
For J = UB(B) To 0 Step -1
    With B(J)
        If .MdNm = "MIde_XEns_CSub" And .MthNm = "WMd_XEns1" Then GoTo Nxt
        If .NeedDlt Then
            If A.Lines(.OldLno, 1) <> .OldCSub Then
                XThw CSub, "OldCSub not expected", _
                    "Md Mth OldLno ExpCSub ActCSub", _
                    Md_Nm(A), .MthNm, .OldLno, .OldCSub, A.Lines(.OldLno, 1)
            End If
            A.DeleteLines .OldLno         '<==
        End If
        If .NeedIns Then
            A.InsertLines .NewLno, .NewCSub
        End If
        '
        If .NeedDlt Or .NeedIns Then
            MsgObj_Prp CSub, "CSub is XEnsured", B(J), "MdNm MthNm NeedDlt OldLno OldCSub NeedIns NewLno NewCSub"
        End If
    End With
Nxt:
Next
End Sub

Private Sub WMd_XEns1a(MdNm$, B() As CSubBrkMth)
Const CSub$ = CMod & "WMd_XEns1a"
If Sz(B) = 0 Then Exit Sub
Dim L1&, L2&
L1 = B(0).OldLno
L2 = B(0).NewLno
Dim J%
For J = 1 To UB(B)
    With B(J)
        If .OldLno > 0 Then
            If L1 > .OldLno Then
                XThw CSub, "CSubBrkMthAy not in sorted order.  That means [Md] has [J] with [Prv-OldLno] > [Cur-OldLno]", _
                    "Md J Prv-OldNo Cur-OldNo", _
                    MdNm, J, L1, .OldLno
            End If
            L1 = .OldLno
        End If
        If L2 > .NewLno Then
            XThw CSub, "[Md] has [J] with [Prv-NewLno] > [Cur-NewLno].  CSubBrkMthAy not in sorted order", _
                "Md J Prv-NewLno Cur-NewLno", _
                MdNm, J, L2, .NewLno
        End If
        L2 = .NewLno
    End With
Next
End Sub

Private Sub WWMd_XEns2(A As CodeModule, B As CSubBrkMd)
Const CSub$ = CMod & "WWMd_XEns2"
With B
    If .NeedDlt Then
        If A.Lines(.OldLno, 1) <> .OldCMod Then
            XThw CSub, "Md CMod is not as expected", _
                "Md LNo OldCMod Exptecd-CMod", _
                Md_Nm(A), .OldLno, A.Lines(.OldLno, 1), .OldCMod
        End If
        A.DeleteLines .OldLno         '<==
    End If
    If .NeedIns Then
        A.InsertLines .NewLno, .NewCMod
    End If
    '
    If .NeedDlt Or .NeedIns Then
        MsgObj_Prp CSub, "CMod is Update", B, "NeedDlt NeedIns OldLno OldCMod NewLno NewCMod"
    End If
End With
End Sub

Private Sub WPj_XEns(A As VBProject)
Dim I
For Each I In Pj_MdAy(A)
   WMd_XEns CvMd(I)
Next
End Sub
