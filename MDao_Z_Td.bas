Attribute VB_Name = "MDao_Z_Td"
Option Compare Binary
Option Explicit

Function CvTd(A) As DAO.TableDef
Set CvTd = A
End Function

Sub Td_XAdd_FdAy(A As DAO.TableDef, FdAy() As DAO.Field2)
Dim I
For Each I In FdAy
    A.Fields.Append I
Next
End Sub

Sub Td_XAdd_IdFld(A As DAO.TableDef)
A.Fields.Append New_Fd(A.Name)
End Sub

Sub Td_XAdd_LngFld(A As DAO.TableDef, FF)
Td_XAdd_FdAy A, ZFdAy(FF, dbLong)
End Sub

Sub Td_XAdd_LngTxt(A As DAO.TableDef, FF)
Td_XAdd_FdAy A, ZFdAy(FF, dbText)
End Sub

Sub Td_XAdd_TimStampFld(A As DAO.TableDef, F$)
A.Fields.Append New_Fd(F, DAO.dbDate, Dft:="Now")
End Sub

Sub TdAddTxtFld(A As DAO.TableDef, FF0, Optional Req As Boolean, Optional Sz As Byte = 255)
Dim F
For Each F In CvNy(FF0)
    A.Fields.Append New_Fd(F, dbText, Req, Sz)
Next
End Sub

Function TdFdScly(A As DAO.TableDef) As String()
Dim N$
N = A.Name & ";"
TdFdScly = Ay_XAdd_Pfx(Itr_Sy_ByMap(A.Fields, "FdScl"), N)
End Function

Function TdFny(A As DAO.TableDef) As String()
TdFny = Fds_Fny(A.Fields)
End Function

Function Td_IsEq(A As DAO.TableDef, B As DAO.TableDef) As Boolean
With A
Select Case True
Case .Name <> B.Name
Case .Attributes <> B.Attributes
Case Not IdxsIsEq(.Indexes, B.Indexes)
Case Not Fds_IsEq(.Fields, B.Fields)
Case Else: Td_IsEq = True
End Select
End With
End Function

Sub Td_IsEq_XAss(A As DAO.TableDef, B As DAO.TableDef)
Dim A1$: A1 = TdStrLines(A)
Dim B1$: B1 = TdStrLines(B)
If A1 <> B1 Then Stop
End Sub

Function Td_Scl$(A As DAO.TableDef)
Td_Scl = Ap_JnSemiColon(A.Name, Val_XAdd_Lbl(A.OpenRecordset.RecordCount, "NRec"), Val_XAdd_Lbl(A.DateCreated, "CrtDte"), Val_XAdd_Lbl(A.LastUpdated, "UpdDte"))
End Function

Function Td_SclLy(A As DAO.TableDef) As String()
Td_SclLy = Ay_XAdd_(Ap_Sy(Td_Scl(A)), TdFdScly(A))
End Function

Function Td_SclLy_XAdd_Pfx(A) As String()
Dim O$(), U&, J&, X
U = UB(A)
If U = -1 Then Exit Function
ReDim O(U)
For Each X In AyNz(A)
    O(J) = IIf(J = 0, "Td;", "Fd;") & X
    J = J + 1
Next
Td_SclLy_XAdd_Pfx = O
End Function

Function TdTyStr$(A As DAO.TableDefAttributeEnum)
TdTyStr = A
End Function

Private Function ZFdAy(FF, T As DAO.DataTypeEnum) As DAO.Field2()
Dim F
For Each F In CvNy(FF)
    PushObj ZFdAy, New_Fd(F, T)
Next
End Function

Private Sub ZZ()
Dim A As Variant
Dim B As DAO.TableDef
Dim C() As DAO.Field2
Dim D$
Dim E As Boolean
Dim F As Byte
Dim G As DAO.TableDefAttributeEnum
CvTd A
Td_XAdd_FdAy B, C
Td_XAdd_IdFld B
Td_XAdd_LngFld B, A
Td_XAdd_LngTxt B, A
Td_XAdd_TimStampFld B, D
TdAddTxtFld B, A, E, F
TdFdScly B
TdFny B
Td_IsEq B, B
Td_IsEq_XAss B, B
Td_Scl B
Td_SclLy B
Td_SclLy_XAdd_Pfx A
TdTyStr G
End Sub

Private Sub Z()
End Sub
