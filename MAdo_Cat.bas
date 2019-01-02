Attribute VB_Name = "MAdo_Cat"
Option Compare Binary
Option Explicit
Function CvCatTbl(A) As ADOX.Table
Set CvCatTbl = A
End Function

Function Cat_Tny(A As Catalog) As String()
Cat_Tny = Itr_Ny(A.Tables)
End Function

Function Fb_Cat(A) As Catalog
Set Fb_Cat = Cn_Cat(Fb_Cn(A))
End Function

Function CvCn(A) As ADODB.Connection
Set CvCn = A
End Function

Sub Cn_XCls(A As ADODB.Connection)
On Error Resume Next
A.Close
End Sub

Function Cn_Cat(A As ADODB.Connection) As Catalog
Dim O As New Catalog
Set O.ActiveConnection = A
Set Cn_Cat = O
End Function

Function Fx_Cat(A) As Catalog
Set Fx_Cat = Cn_Cat(Fx_Cn(A))
End Function
