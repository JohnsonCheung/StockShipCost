Attribute VB_Name = "MCrypto__Fun"
Option Compare Binary
Option Explicit
Sub XXXX()
Dim X As mscorlib.Object
Dim T
Set X = New mscorlib.Object
Stop
Set T = X.getType()
Stop
End Sub
Private Sub Z_AsmAy_NotWrk()
Dim O
O = AsmAy_NotWrk
Stop
End Sub
Property Get AsmAy_NotWrk() As Object()
' function Get-Assemblies { [System.AppDomain]::CurrentDomain.GetAssemblies()
Dim AppDomain, CurDomain
AppDomain = CreateObject("System.AppDomain")
Set CurDomain = AppDomain.CurrentDomain
Stop
End Property
Sub YY()
Static A As excel.Application
Set A = New excel.Application
Stop
End Sub
Sub XX()
Dim A As excel.Application
Set A = GetObject(, "Excel.Application")
A.Workbooks.Add
A.Visible = False 'Must have workbook open to allow Visible has effect
Dim B As excel.Application
Set B = GetObject(, "Excel.Application")
B.Workbooks.Add
B.Visible = False 'Must have workbook open to allow Visible has effect
Debug.Print ObjPtr(A), ObjPtr(B)
Stop
Stop
End Sub
Function ToBase64String(rabyt)

  'Ref: http://stackoverflow.com/questions/1118947/converting-binary-file-to-base64-string
  With CreateObject("MSXML2.DOMDocument")
    .LoadXML "<root />"
    .DocumentElement.DataType = "bin.base64"
    .DocumentElement.nodeTypedValue = rabyt
    ToBase64String = Replace(.DocumentElement.text, vbLf, "")
  End With
End Function

Function ToHexString(rabyt)

  'Ref: http://stackoverflow.com/questions/1118947/converting-binary-file-to-base64-string
  With CreateObject("MSXML2.DOMDocument")
    .LoadXML "<root />"
    .DocumentElement.DataType = "bin.Hex"
    .DocumentElement.nodeTypedValue = rabyt
    ToHexString = Replace(.DocumentElement.text, vbLf, "")
  End With
End Function

Sub to_SHA256()
'Requires a reference to mscorlib 4.0 64-bit, which is part of the .Net Framework 4.0
GoTo Tst1
Exit Sub
Tst1:
    Dim A() As Byte
    Dim text As Object
    Dim SHA256 As Object
        A = CreateObject("System.Text.UTF8Encoding").GetBytes_4("abcd")
        Set text = CreateObject("System.Text.UTF8Encoding")
        Set SHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
        
        If True Then
            Dim Bytes
            Dim Hash
            Bytes = text.GetBytes_4("mypassword")
            Hash = SHA256.ComputeHash_2((Bytes)) ' Single brackket quote is not OK
            Debug.Print ToHexString(Hash)
        Else
            Debug.Print ToHexString(SHA256.ComputeHash_2((text.GetBytes_4("mypassword"))))
        End If
        Stop
    XShw_Dbg
    Stop
    Return
End Sub

'64-bit MS Access VBA code to calculate an SHA-512 or SHA-256 hash in VBA.  This requires a VBA reference to the .Net Framework 4.0 mscorlib.dll.  The hashed strings are calculated using calls to encryption methods built into mscorlib.dll.  The calculated hash strings are the same values as those calculated with jsSHA, a Javascript SHA implementation (see https://caligatio.github.io/jsSHA/ for an online calculator and the jsSHA code).
'The mscorlib.dll is intended for .Net Framework managed applications, but the stackoverflow.com post showed how it could be used with MS Access VBA.  This technique is not documented anywhere in MS Access documentation that I could find, so the stackoverflow.com post was very helpful in this regard.
Sub to_SHA512()
'Requires a reference to mscorlib 4.0 64-bit
Dim text As Object
Dim SHA512 As Object
Dim SHA256 As Object

Set text = CreateObject("System.Text.UTF8Encoding")

Set SHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")
Set SHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")

Debug.Print ToBase64String(SHA512.ComputeHash_2((text.GetBytes_4("mypassword"))))
Debug.Print ToHexString(SHA512.ComputeHash_2((text.GetBytes_4("mypassword"))))
End Sub

Sub XXX()
Dim X
Set X = CreateObject("System.Collections.ArrayList")
X.Add 1
Dim J%
For J = 1 To 1000
    X.Add J
Next
Dim I
For Each I In X
    Debug.Print I
Next
Stop
End Sub
Private Sub ZZ()
Dim A As Variant
ToBase64String A
ToHexString A
to_SHA256
to_SHA512
End Sub

Private Sub Z()
End Sub
