Attribute VB_Name = "MXls__Samp"
Option Compare Binary
Option Explicit
Property Get SampLo_XVis() As ListObject
Set SampLo_XVis = Lo_XVis(SampLo)
End Property

Property Get SampLo() As ListObject
Set SampLo = Rg_Lo(Sq_XPut_AtCell(SampSqWithHdr, New_A1))
End Property

Property Get SampLoFmtr() As String()
Dim O$()
Push O, "Lo Nam ABC"
Push O, "Lo Fld A B C D E F G"
Push O, "Lo Hid B C X"
Push O, "Wdt 10 A B X"
Push O, "Wdt 20 D C C"
Push O, "Wdt 3000 E F G C"
Push O, "Fmt #,## A B C"
Push O, "Fmt #,##.## D E"
Push O, "Lvl 2 A C"
Push O, "Bdr Left A"
Push O, "Bdr Right G"
Push O, "Bdr Col F"
Push O, "Tot Sum A B"
Push O, "Tot Cnt C"
Push O, "Tot Avg D"
Push O, "Tit A abc | sdf"
Push O, "Tit B abc | sdkf | sdfdf"
Push O, "Cor 12345 A B"
Push O, "Fml F A + B"
Push O, "Fml C A * 2"
Push O, "Lbl A lksd flks dfj"
Push O, "Lbl B lsdkf lksdf klsdj f"
SampLoFmtr = O
End Property

Property Get Samp_LoFmtrTp() As String()
Dim O$()
PushI O, "Lo  Nm     *Nm"
PushI O, "Lo  Fld    *Fld.."
PushI O, "Align Left   *Fld.."
PushI O, "Align Right  *Fld.."
PushI O, "Align Center *Fld.."
PushI O, "Bdr Left   *Fld.."
PushI O, "Bdr Right  *Fld.."
PushI O, "Bdr Col    *Fld.."
PushI O, "Tot Sum    *Fld.."
PushI O, "Tot Avg    *Fld.."
PushI O, "Tot Cnt    *Fld.."
PushI O, "Fmt *Fmt   *Fld.."
PushI O, "Wdt *Wdt   *Fld.."
PushI O, "Lvl *Lvl   *Fld.."
PushI O, "Cor *Cor   *Fld.."
PushI O, "Fml *Fld   *Formula"
PushI O, "Bet *Fld   *Fld1 *Fld2"
PushI O, "Tit *Fld   *Tit"
PushI O, "Lbl *Fld   *Lbl"
Samp_LoFmtrTp = O
End Property

Property Get SampSq() As Variant()
Const NR% = 10
Const NC% = 10
Dim O(), R%, C%
ReDim O(1 To NR, 1 To NC)
SampSq = O
For R = 1 To NR
    For C = 1 To NC
        O(R, C) = R * 1000 + C
    Next
Next
SampSq = O
End Property

Property Get SampSqHdr() As Variant()
Const NC% = 10
Dim J%
For J = 0 To NC - 1
    PushI SampSqHdr, Chr(Asc("A") + J)
Next
End Property

Property Get SampSqWithHdr() As Variant()
SampSqWithHdr = Sq_Ins_Dr(SampSq, SampSqHdr)
End Property

Property Get SampWs() As Worksheet
Dim O As Worksheet
Set O = New_Ws
Drs_Lo Samp_Drs, Ws_RC(O, 2, 2)
Set SampWs = O
Ws_XVis O
End Property
