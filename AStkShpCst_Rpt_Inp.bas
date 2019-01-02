Attribute VB_Name = "AStkShpCst_Rpt_Inp"
Option Compare Binary
Option Explicit
Property Get IFxAy() As String()
IFxAy = Ap_Sy(IFxMB52, IFxUOM, IFxZHT1)
End Property

Property Get IFxMB52$()
IFxMB52 = Pnm_Ffn("MB52")
End Property

Property Get IFxUOM$()
IFxUOM = Pnm_Ffn("UOM")
End Property

Property Get IFxZHT1$()
IFxZHT1 = Pnm_Ffn("ZHT1")
End Property

Sub OpnIMB52(): Fx_XBrw IFxMB52: End Sub

Sub OpnIUOM(): Fx_XBrw IFxUOM: End Sub

Sub OpnIZHT1(): Fx_XBrw IFxZHT1: End Sub

Property Get IZHT1Fny() As String()
Ay_XDmp Dbt_Fny(W, ">ZHT1")
End Property

Property Get PmStkDte() As Date
Dim A$
A = Mid(Pnm_Val("MB52Fn"), 6, 10)
PmStkDte = CDate(A)
End Property

Property Get PmStkYYMD$()
PmStkYYMD = Format(PmStkDte, "YYYY-MM-DD")
End Property
