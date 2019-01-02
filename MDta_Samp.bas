Attribute VB_Name = "MDta_Samp"
Option Compare Binary
Option Explicit
Property Get Samp_Dr1() As Variant()
Samp_Dr1 = Array(1, 2, 3)
End Property

Property Get Samp_Dr2() As Variant()
Samp_Dr2 = Array(2, 3, 4)
End Property

Property Get Samp_Dr3() As Variant()
Samp_Dr3 = Array(3, 4, 5)
End Property

Property Get Samp_Dr4() As Variant()
Samp_Dr4 = Array(43, 44, 45)
End Property

Property Get Samp_Dr5() As Variant()
Samp_Dr5 = Array(53, 54, 55)
End Property

Property Get Samp_Dr6() As Variant()
Samp_Dr6 = Array(63, 64, 65)
End Property

Property Get Samp_Drs1() As Drs
Set Samp_Drs1 = New_Drs("A B C", Samp_Dry1)
End Property

Property Get Samp_Drs2() As Drs
Set Samp_Drs2 = New_Drs("A B C", Samp_Dry2)
End Property

Property Get Samp_Drs() As Drs
Set Samp_Drs = New_Drs("A B C D E G H I J K", Samp_Dry)
End Property

Property Get Samp_DRs_Fny() As String()
Samp_DRs_Fny = Ssl_Sy("A B C D E F G")
End Property

Property Get Samp_Dry1() As Variant()
Samp_Dry1 = Array(Samp_Dr1, Samp_Dr2, Samp_Dr3)
End Property

Property Get Samp_Dry2() As Variant()
Samp_Dry2 = Array(Samp_Dr3, Samp_Dr4, Samp_Dr5)
End Property

Property Get Samp_Dry() As Variant()
PushI Samp_Dry, Array("A", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "A"))
PushI Samp_Dry, Array("B", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "B"))
PushI Samp_Dry, Array("C", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "C"))
PushI Samp_Dry, Array("D", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "D"))
PushI Samp_Dry, Array("E", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "E"))
PushI Samp_Dry, Array("F", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "F"))
PushI Samp_Dry, Array("G", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "G"))
End Property

Property Get Samp_Ds() As Ds
Set Samp_Ds = Ds(Ap_DtAy(Samp_Dt1, Samp_Dt2), "Samp_Ds")
End Property

Property Get Samp_Dt1() As Dt
Set Samp_Dt1 = New_Dt("Samp_Dt1", "A B C", Samp_Dry1)
End Property

Property Get Samp_Dt2() As Dt
Set Samp_Dt2 = New_Dt("Samp_Dt2", "A B C", Samp_Dry2)
End Property
