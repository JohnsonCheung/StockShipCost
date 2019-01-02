Attribute VB_Name = "MXls_Z_Xls"
Option Compare Binary
Option Explicit
Function Xls_Wb_ByFx(A As excel.Application, Fx) As Workbook
Dim W As Workbook
For Each W In A.Workbooks
    If Wb_Fx(W) = Fx Then
        Set Xls_Wb_ByFx = W
        Exit Function
    End If
Next
End Function
Function Xls(Optional Vis As Boolean) As excel.Application
Static Y As excel.Application
On Error GoTo XX
Dim J%, A$
Beg:
A = Y.Name
Set Xls = Y
Exit Function
XX:
    J = J + 1
    If J > 10 Then Stop
    Set Y = New excel.Application
    GoTo Beg
End Function

Function XlsAddIn(A As excel.Application, FxaNm) As excel.AddIn
Dim I As excel.AddIn
For Each I In A.AddIns
    If Str_IsEq(I.Name, FxaNm & ".xlam") Then Set XlsAddIn = I
Next
End Function

Property Get CurXls() As excel.Application
Static X As excel.Application
If IsNothing(X) Then Set X = excel.Application
Set CurXls = X
End Property

Function XlsHasAddInFn(A As excel.Application, AddInFn) As Boolean
Dim I As excel.AddIn
Dim N$: N = UCase(AddInFn)
For Each I In A.AddIns
    If UCase(I.Name) = N Then XlsHasAddInFn = True: Exit Function
Next
End Function
Sub CurXls_XQuit()
Xls_XQuit CurXls
End Sub
Sub Xls_XQuit(A As excel.Application)
Itr_XDo A.Workbooks, "Wb_XCls_NoSav"
A.Quit
Set A = Nothing
End Sub

Sub Xls_XVis(A As excel.Application)
If Not A.Visible Then A.Visible = True
End Sub
