Attribute VB_Name = "MVb_Str_Mch"
Option Compare Binary
Option Explicit
Type PatnRslt
    Patn As String
    Rslt As Variant
End Type

Function StrDicMch(A, PatnRsltDic As Dictionary) As PatnRslt
Dim Patn
For Each Patn In PatnRsltDic
    If XMch_Patn(A, Patn) Then
        With StrDicMch
            .Rslt = PatnRsltDic(Patn)
            .Patn = Patn
        End With
        Exit Function
    End If
Next
End Function

Private Sub Z_XMch_Patn()
Dim A$, Patn$
Ept = True: A = "AA": Patn = "AA": GoSub Tst
Ept = True: A = "AA": Patn = "^AA$": GoSub Tst
Exit Sub
Tst:
    Act = XMch_Patn(A, Patn)
    C
    Return
End Sub

Function XMch_Patn(S, Patn) As Boolean
Static Re As New RegExp
Re.Pattern = Patn
XMch_Patn = Re.Test(S)
End Function


Private Sub Z()
Z_XMch_Patn
MVb_Str_Mch:
End Sub
