Attribute VB_Name = "AStkShpCst_Rpt"
Option Compare Binary
Option Explicit
Dim Dbg1 As Boolean
Const LnkColVblzZHT1$ = _
" ZHT1   Txt Brand  |" & _
" RateSc Dbl Amount |" & _
" VdtFm  Txt [Valid From]  |" & _
" VdtTo  Txt [Valid to]"

Const LnkColVblzUom$ = _
 "Sku    Txt Material |" & _
 "Des    Txt [Material Description] |" & _
 "Sc_U   Txt;Dbl SC |" & _
 "StkUom Txt [Base Unit of Measure] |" & _
 "Topaz  Txt [Topaz Code] |" & _
 "ProdH  Txt [Product hierarchy]"
 
Const LnkColVblzMB52$ = _
    " Sku    Txt Material |" & _
    " Whs    Txt Plant    |" & _
    " QInsp  Dbl [In Quality Insp#]|" & _
    " QUnRes Dbl Unrestricted|" & _
    " QBlk   Dbl Blocked"

Sub DocUOM()
'InpX: [>UOM]     Material [Base Unit of Measure] [Material Description] [Unit per case]
'Oup : UOM        Sku      SkuUOM                 Des                    Sc_U

'Note on [Sales text.xls]
'Col  Xls Title            FldName     Means
'F    Base Unit of Measure SkuUOM      either COL (bottle) or PCE (set)
'J    Unit per case        Sc_U        how many unit per AC
'K    SC                   SC_U        how many unit per SC   ('no need)
'L    COL per case         AC_B        how many bottle per AC
'-----
'Letter meaning
'B = Bottle
'AC = act case
'SC = standard case
'U = Unit  (Bottle(COL) or Set (PCE))

' "SC              as SC_U," & _  no need
' "[COL per case]  as AC_B," & _ no need
End Sub

Sub XRptImp()
'Create 5-Imp-Table [#I*] from 5-lnk-table [>*]
WImp ">UOM", LnkColVblzUom
WImp ">ZHT18601", LnkColVblzZHT1
WImp ">ZHT18701", LnkColVblzZHT1
WImp ">MB52", LnkColVblzMB52
End Sub

Sub XRptLnk()
Dim A$(), B$(), C$(), D$()
A = Wt_XLnk_Fx(">UOM", IFxUOM)
B = Wt_XLnk_Fx(">MB52", IFxMB52)
C = Wt_XLnk_Fx(">ZHT18601", IFxZHT1, "8601")
D = Wt_XLnk_Fx(">ZHT18701", IFxZHT1, "8701")
Er_XHalt Ap_Sy(A, B, C, D)
A = Wt_XChk_Col_ByLnkColVbl(">UOM", LnkColVblzUom)
B = Wt_XChk_Col_ByLnkColVbl(">MB52", LnkColVblzMB52)
C = Wt_XChk_Col_ByLnkColVbl(">ZHT18601", LnkColVblzZHT1)
D = Wt_XChk_Col_ByLnkColVbl(">ZHT18701", LnkColVblzZHT1)
Er_XHalt Ap_Sy(A, B, C, D)
End Sub

Sub OMain()
WDrp "@Main"
WRun "Select Distinct Whs,Sku,Sum(QUnRes+QBlk+QInsp) As OH into [@Main] from [#IMB52] Group by Whs,Sku"

'Des StkUom Sc_U OH_Sc
WRun "Alter Table [@Main] Add Column Des Text(255), StkUom Text(10),Sc_U Int, OH_Sc Double"
WRun "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.Sc_U = a.Sc_U,x.Des=a.Des,x.StkUom=a.StkUom"
WRun "Update [@Main] set OH_Sc=OH/Sc_U where Sc_U>0"

'Stream ProdH F2 M32 M35 M37 Topaz ZHT1 RateSc Z2 Z5 Z7
WRun "Alter Table [@Main] add column Stream Text(10), Topaz Text(20), ProdH text(7), F2 Text(2), M32 text(2), M35 text(5), M37 text(7), ZHT1 Text(7), Z2 text(2), Z5 text(5), Z7 text(7), RateSc Currency, Amt Currency"

'ProdH Topaz
WRun "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.ProdH=a.ProdH,x.Topaz=a.Topaz"

'F2 M32 M35 M37
WRun "Update [@Main] set F2=Left(ProdH,2),M32=Mid(ProdH,3,2),M35=Mid(ProdH,3,5),M37=Mid(ProdH,3,7)"

'ZHT1 RateSc
WRun "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M37=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
WRun "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M35=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
WRun "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M32=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"

'Stream
WRun "Update [@Main] set Stream=IIf(Left(Topaz,3)='UDV','Diageo','MH')"

'Z2 Z5 Z7
WRun "Update [@Main] Set Z2=Left(ZHT1,2), Z5=Left(ZHT1,5), Z7=Left(ZHT1,7) where not ZHT1 is null"

'Amt
WRun "Update [@Main] Set Amt = RateSc * OH_Sc where RateSc is not null"
End Sub

Sub ORate()
'VdtFm & VdtTo format DD.MM.YYYY
'1: #IZHT1 VdtFm VdtTo L3 RateSc
'2: #IUom     SKu Sc_U
'O: @Rate  ZHT1 RateSc
WDrp "#Cpy1 #Cpy2 #Cpy @Rate"
WRun "Select '8701' as Whs,x.* into [#Cpy1] from [#IZHT18701] x"
WRun "Select '8601' as Whs,x.* into [#Cpy2] from [#IZHT18601] x"

WRun "Select * into [#Cpy] from [#Cpy1] where False"
WRun "Insert into [#Cpy] select * from [#Cpy1]"
WRun "Insert into [#Cpy] select * from [#Cpy2]"

WRun "Alter Table [#Cpy] Add Column VdtFmDte Date,VdtToDte Date,IsCur YesNo"
WRun "Update [#Cpy] Set" & _
" VdtFmDte = DateSerial(RIGHT(VdtFm,4),MID(VdtFm,4,2),LEFT(VdtFm,2))," & _
" VdtToDte = DateSerial(RIGHT(VdtTo,4),MID(VdtTo,4,2),LEFT(VdtTo,2))"
WRun "Update [#Cpy] set IsCur = true where Now between VdtFmDte and VdtToDte"

WRun "Select Whs,ZHT1,RateSc into [@Rate] from [#Cpy]"
WDrp "#Cpy #Cpy1 #Cpy2"
End Sub

Sub XRptOup()
Qnm_XSet_MainMsg "@Rate": ORate
Qnm_XSet_MainMsg "@Main": OMain
End Sub

Property Get OupFx$()
Dim A$, B$
A = OupPth & QQ_Fmt("? ?.xlsx", Apn, Mid(Pnm_Val("MB52Fn"), 6, 10))
B = Ffn_NxtFfn(A)
OupFx = B
End Property

Property Get OupPth$()
OupPth = Pth_XEns(CurDb_Pth & "Output\")
End Property
Sub XIni()
WIni
End Sub
Sub XRptGen()
WGen OupFx
End Sub
Sub XRpt()
XClr_MainMsg
XIni
XRptLnk
XRptEr
XRptImp
XRptOup
XRptGen
End Sub
