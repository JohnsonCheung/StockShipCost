Attribute VB_Name = "MXls_Z_Xls_AddIn"
Option Compare Binary
Option Explicit

Function XlsAddInDrs(A As excel.Application) As Drs
Set XlsAddInDrs = ItoDrs(A.AddIns, "Name FullName Installed IsOpen ProgId CLSID")
End Function

Sub XlsAddInDmp(A As excel.Application)
Drs_XDmp XlsAddInDrs(A)
End Sub

Sub CurXlsAddInDmp()
XlsAddInDmp CurXls
End Sub

Property Get CurXlsAddInWs() As Worksheet
Set CurXlsAddInWs = Ws_XVis(Drs_Ws(XlsAddInDrs(CurXls)))
End Property

