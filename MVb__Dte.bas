Attribute VB_Name = "MVb__Dte"
Option Compare Binary
Option Explicit
Property Get CurM() As Byte
CurM = Month(Now)
End Property

Function M_NxtM(M As Byte) As Byte
M_NxtM = IIf(M = 12, 1, M + 1)
End Function

Function M_PrvM(M As Byte) As Byte
M_PrvM = IIf(M = 1, 12, M - 1)
End Function

Property Get NowDTim$()
NowDTim = Dte_DTim(Now)
End Property

Property Get NowStr$()
NowStr = NowDTim
End Property

Function Dte_DTim$(A As Date)
Dte_DTim = Format(A, "YYYY-MM-DD HH:MM:SS")
End Function

Function Dte_FstDayOfMth(A As Date) As Date
Dte_FstDayOfMth = DateSerial(Year(A), Month(A), 1)
End Function

Function Dte_FstDteOfMth(A As Date) As Date
Dte_FstDteOfMth = DateSerial(Year(A), Month(A), 1)
End Function

Function DteIsVdt(A$) As Boolean
On Error Resume Next
DteIsVdt = Format(CDate(A), "YYYY-MM-DD") = A
End Function

Function Dte_LasDayOfMth(A As Date) As Date
Dte_LasDayOfMth = Dte_PrvDay(Dte_FstDteOfMth(Dte_NxtMth(A)))
End Function

Function Dte_NxtMth(A As Date) As Date
Dte_NxtMth = DateTime.DateAdd("M", 1, A)
End Function

Function Dte_PrvDay(A As Date) As Date
Dte_PrvDay = DateAdd("D", -1, A)
End Function

Function Dte_YYMM$(A As Date)
Dte_YYMM = Right(Year(A), 2) & Format(Month(A), "00")
End Function

Function YYMM_FstDte(A) As Date
YYMM_FstDte = DateSerial(Left(A, 2), Mid(A, 3, 2), 1)
End Function

Function YYYYMMDD_IsVdt(A) As Boolean
On Error Resume Next
YYYYMMDD_IsVdt = Format(CDate(A), "YYYY-MM-DD") = A
End Function

Function YM_FstDte(Y As Byte, M As Byte) As Date
YM_FstDte = DateSerial(2000 + Y, M, 1)
End Function

Function YM_LasDte(Y As Byte, M As Byte) As Date
YM_LasDte = Dte_NxtMth(YM_FstDte(Y, M))
End Function

Function YM_YofNxtM(Y As Byte, M As Byte) As Byte
YM_YofNxtM = IIf(M = 12, Y + 1, Y)
End Function

Function YM_YofPrvM(Y As Byte, M As Byte) As Byte
YM_YofPrvM = IIf(M = 1, Y - 1, Y)
End Function

Property Get CurY() As Byte
CurY = CurYY - 2000
End Property

Property Get CurYY%()
CurYY = Year(Now)
End Property
