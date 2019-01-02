Attribute VB_Name = "MSql_QUpd"
Option Compare Binary
Option Explicit
Private X As New Sql_Shared
Function UpdSqlFmt$(T, Sk$(), Fny$(), Dr)
If Sz(Sk) = 0 Then Stop
Dim Upd$, Set_$, Wh$: GoSub X_Upd_Set_Wh
UpdSqlFmt = Upd & Set_ & Wh
Exit Function
X_Upd_Set_Wh:
    Dim Fny1$(), Dr1(), SkVV(): GoSub X_Fny1_Dr1_SkVV
    Upd = "Update [" & T & "]"
    Set_ = FnyVy_SetSqp(Fny1, Dr1)
    Wh = QQWh_FnyEqVy(Sk, SkVV)
    Return
X_Ay:
    Dim L$(), R$()
    L = FnyAlignQuote(Fny)
    R = Vy_XQuote_Sql(Dr)
    Return
X_Fny1_Dr1_SkVV:
    Dim Ski, J%, IxAy%(), I%
    For Each Ski In Sk
        I = Ay_Ix(Fny, Ski)
        If I = -1 Then Stop
        Push IxAy, I
        Push SkVV, Dr(I)    '<====
    Next
    Dim F
    For Each F In Fny
        If Not Ay_XHas(IxAy, J) Then
            Push Fny1, F        '<===
            Push Dr1, Dr(J)     '<===
        End If
        J = J + 1
    Next
    Return
End Function

Private Sub Z_UpdSqlFmt()
Dim T$, Sk$(), Fny$(), Dr
T = "A"
Sk = Lin_TermAy("X Y")
Fny = Lin_TermAy("X Y A B C")
Dr = Array(1, 2, 3, 4, 5)
Ept = "Update [A]" & _
vbCrLf & "  Set" & _
vbCrLf & "    A = 3 ," & _
vbCrLf & "    B = 4 ," & _
vbCrLf & "    C = 5 " & _
vbCrLf & "  Where" & _
vbCrLf & "    X = 1 And" & _
vbCrLf & "    Y = 2 "
GoSub Tst

T = "A"
Sk = Lin_TermAy("[A 1] B CD")
Fny = Lin_TermAy("X Y B Z CD [A 1]")
Dr = Array(1, 2, 3, 4, "XX", #1/2/2018 12:34:00 PM#)
Ept = "Update [A]" & _
vbCrLf & "  Set" & _
vbCrLf & "    X = 1 ," & _
vbCrLf & "    Y = 2 ," & _
vbCrLf & "    Z = 4 " & _
vbCrLf & "  Where" & _
vbCrLf & "    [A 1] = #2018-01-02 12:34:00# And" & _
vbCrLf & "    B     = 3                     And" & _
vbCrLf & "    CD    = 'XX'                  "
GoSub Tst
Exit Sub
Tst:
    Act = UpdSql(T, Sk, Fny, Dr)
    C
    Return
End Sub

Function QAddCol$(T, Fny0, F As Drs, E As Dictionary)
Dim O$(), Fld
For Each Fld In CvNy(Fny0)
'    PushI O, Fld & " " & QAddCol1(Fld, F, E)
Next
QAddCol = QQ_Fmt("Alter Table [?] add column ?", T, JnComma(O))
End Function

Private Function QAddCol1$(F, EF As EF)
QAddCol1 = Fd_SqlTy(New_Fd_EF(F, "", EF))
End Function

Function Tbl_CrtPkSql$(T)
Tbl_CrtPkSql = QQ_Fmt("Create Index PrimaryKey on [?] (?Id) with Primary", T, T)
End Function

Function TblSkFF_CrtSkSql$(T, SkFF)
TblSkFF_CrtSkSql = QQ_Fmt("Create unique Index SecondaryKey on [?] (?)", T, FF_JnComma(SkFF))
End Function

Function CrtTblSql$(T, FldList$)
CrtTblSql = QQ_Fmt("Create Table [?] (?)", T, FldList)
End Function

Function DrpFldSql$(T, F)
DrpFldSql = QQ_Fmt("Alter Table [?] drop column [?]", T, F)
End Function

Function DrpTblSql$(T)
DrpTblSql = "Drop Table [" & T & "]"
End Function

Function DrsCrtTblSql$(A As Drs, T)
Dim F, J%, Dry(), O$()
Dry = A.Dry
For Each F In A.Fny
    PushI O, F & " " & DryColSqlTy(Dry, J)
    J = J + 1
Next
DrsCrtTblSql = CrtTblSql(T, JnComma(O))
End Function

Function InsDrSql$(T, Fny0, Dr)
InsDrSql = QQ_Fmt("Insert into [?] (?) values(?)", T, JnComma(CvNy(Fny0)), JnComma(Vy_XQuote_Sql(Dr)))
End Function

Function InsSql$(T, Fny$(), Dr)
Dim A$, B$
A = JnComma(Fny)
B = JnComma(AyMap_Sy(Dr, "Var_XQuote_Sql"))
InsSql = QQ_Fmt("Insert Into [?] (?) Values(?)", T, A, B)
End Function

Function SelTblSql$(T)
SelTblSql = "Select * from [" & T & "]"
End Function

Function SelTblWhSql$(T, WhBExpr$)
SelTblWhSql = SelTblSql(T) & QQWh_BExpr(WhBExpr)
End Function

Function SelTFSql$(T, F)
SelTFSql = QQ_Fmt("Select [?] from [?]", F, T)
End Function

Function UpdSql$(T, Sk$(), Fny$(), Dr)
If Sz(Sk) = 0 Then Stop
Dim Upd$, Set_$, Wh$: GoSub X_Upd_Set_Wh
UpdSql = Upd & Set_ & Wh
Exit Function
X_Upd_Set_Wh:
    Dim Fny1$(), Dr1(), SkVV(): GoSub X_Fny1_Dr1_SkVV
    Upd = "Update [" & T & "]"
    Set_ = FnyVy_SetSqp(Fny1, Dr1)
    Wh = QQWh_FnyEqVy(Sk, SkVV)
    Return
X_Ay:
    Dim L$(), R$()
    L = FnyAlignQuote(Fny)
    R = Vy_XQuote_Sql(Dr)
    Return
X_Fny1_Dr1_SkVV:
    Dim Ski, J%, IxAy%(), I%
    For Each Ski In Sk
        I = Ay_Ix(Fny, Ski)
        If I = -1 Then Stop
        Push IxAy, I
        Push SkVV, Dr(I)    '<====
    Next
    Dim F
    For Each F In Fny
        If Not Ay_XHas(IxAy, J) Then
            Push Fny1, F        '<===
            Push Dr1, Dr(J)     '<===
        End If
        J = J + 1
    Next
    Return
End Function



Function FnyVy_SetSqp$(Fny$(), Vy())
Dim A$: GoSub X_A
FnyVy_SetSqp = "  Set " & A
Exit Function
X_A:
    Dim L$(): L = Ay_XQuote_SqBkt_IfNeed(Fny)
    Dim R$(): R = Vy_XQuote_Sql(Vy)
    Dim J%, O$()
    For J = 0 To UB(L)
        Push O, L(J) & " = " & R(J)
    Next
    A = JnComma(O)
    Return
End Function
Function FnyAlignQuote(Fny$()) As String()
FnyAlignQuote = Ay_XAlign_L(Ay_XQuote_SqBkt_IfNeed(Fny))
End Function




Private Sub Z()
Z_UpdSqlFmt
MSql_Upd:
End Sub
