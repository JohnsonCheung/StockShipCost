Attribute VB_Name = "MTp_Sq__Samp"
Option Compare Binary
Option Explicit
Property Get Samp_SqTp$()
Dim O$()
PushI O, "-- Rmk: -- is remark"
PushI O, "-- %XX: is prmDicLin"
PushI O, "-- %?XX: is switchPrm, it value must be 0 or 1"
PushI O, "-- ?XX: is switch line"
PushI O, "-- SwitchLin: is ?XXX [OR|AND|EQ|NE] [SwPrm_OR_AND|SwPrm_EQ_NE]"
PushI O, "-- SwPrm_OR_AND: SwTerm .."
PushI O, "-- SwPrm_EQ_NE:  SwEQ_NE_T1 SwEQ_NE_T2"
PushI O, "-- SwEQ_NE_T1:"
PushI O, "-- SwEQ_NE_T2:"
PushI O, "-- SwTerm:     ?XX|%?XX     -- if %?XX, its value only 1 or 0 is allowed"
PushI O, "-- Only one gp of %XX:"
PushI O, "-- Only one gp of ?XX:"
PushI O, "-- All other gp is sql-statement or sql-statements"
PushI O, "-- sql-statments: Drp xxx xxx"
PushI O, "-- sql-statment: [sel|selDis|upd|into|fm|whBetStr|whBetNbr|whInStrLis|whInNbrLis|andInNbrLis|andInStrLis|gp|jn|left|expr]"
PushI O, "-- optional: Whxxx and Andxxx can have ?-pfx becomes: ?Whxxx and ?Andxxx.  The line will become empty"
PushI O, "=============================================="
PushI O, "Drp Tx TxMbr MbrDta Div Sto Crd Cnt Oup MbrWs"
PushI O, "============================================="
PushI O, "-- @? means switch, value must be 0 or 1"
PushI O, "@?BrkMbr 0"
PushI O, "@?BrkMbr 0"
PushI O, "@?BrkMbr 0"
PushI O, "@?BrkSto 0"
PushI O, "@?BrkCrd 0"
PushI O, "@?BrkDiv 0"
PushI O, "-- %XXX means txt and optional, allow, blank"
PushI O, "@SumLvl  Y"
PushI O, "@?MbrEmail 1"
PushI O, "@?MbrNm    1"
PushI O, "@?MbrPhone 1"
PushI O, "@?MbrAdr   1"
PushI O, "-- %% mean compulasary"
PushI O, "@%DteFm 20170101"
PushI O, "@%DteTo 20170131"
PushI O, "@LisDiv 1 2"
PushI O, "@LisSto"
PushI O, "@LisCrd"
PushI O, "@CrdExpr ..."
PushI O, "@CrdExpr ..."
PushI O, "@CrdExpr ..."
PushI O, "============================================"
PushI O, "-- EQ & NE t1 only TxtPm is allowed"
PushI O, "--         t2 allow TxtPm, *BLANK, and other text"
PushI O, "?LvlY    EQ %SumLvl Y"
PushI O, "?LvlM    EQ %SumLvl M"
PushI O, "?LvlW    EQ %SumLvl W"
PushI O, "?LvlD    EQ %SumLvl D"
PushI O, "?Y       OR ?LvlD ?LvlW ?LvlM ?LvlY"
PushI O, "?M       OR ?LvlD ?LvlW ?LvlM"
PushI O, "?W       OR ?LvlD ?LvlW"
PushI O, "?D       OR ?LvlD"
PushI O, "?Dte     OR ?LvlD"
PushI O, "?Mbr     OR %?BrkMbr"
PushI O, "?MbrCnt  OR %?BrkMbr"
PushI O, "?Div     OR %?BrkDiv"
PushI O, "?Sto     OR %?BrkSto"
PushI O, "?Crd     OR %?BrkCrd"
PushI O, "?#SEL#Div NE %LisDiv *blank"
PushI O, "?#SEL#Sto NE %LisSto *blank"
PushI O, "?#SEL#Crd NE %LisCrd *blank"
PushI O, "============================================= #Tx"
PushI O, "sel  ?Crd ?Mbr??Div ?Sto ?Y ?M ?W ?WD ?D ?Dte Amt Qty Cnt"
PushI O, "into #Tx"
PushI O, "fm   SalesHistory"
PushI O, "wh   bet str    %%DteFm %%DteTo"
PushI O, "?and in  strlis Div %LisDiv"
PushI O, "?and in  strlis Sto %LisSto"
PushI O, "?and in  nbrlis Crd %LisCrd"
PushI O, "?gp  ?Crd ?Mbr ?Div ?Sto ?Crd ?Y ?M ?W ?WD ?D ?Dte"
PushI O, "$Crd %CrdExpr"
PushI O, "$Mbr JCMCode"
PushI O, "$Sto"
PushI O, "$Y"
PushI O, "$M"
PushI O, "$W"
PushI O, "$WD"
PushI O, "$D"
PushI O, "$Dte"
PushI O, "$Amt Sum(SHAmount)"
PushI O, "$Qty Sum(SHQty)"
PushI O, "$Cnt Count(SHInvoice+SHSDate+SHRef)"
PushI O, "============================================= #TxMbr"
PushI O, "selDis  Mbr"
PushI O, "fm      #Tx"
PushI O, "into    #TxMbr"
PushI O, "============================================= #MbrDta"
PushI O, "sel   Mbr Age Sex Sts Dist Area"
PushI O, "fm    #TxMbr x"
PushI O, "jn    JCMMember a on x.Mbr = a.JCMMCode"
PushI O, "into  #MbrDta"
PushI O, "$Mbr  x.Mbr"
PushI O, "$Age  DATEDIFF(YEAR,CONVERT(DATETIME ,x.JCMDOB,112),GETDATE())"
PushI O, "$Sex  a.JCMSex"
PushI O, "$Sts  a.JCMStatus"
PushI O, "$Dist a.JCMDist"
PushI O, "$Area a.JCMArea"
PushI O, "==-=========================================== #Div"
PushI O, "?sel Div DivNm DivSeq DivSts"
PushI O, "fm   Division"
PushI O, "into #Div"
PushI O, "?wh in strLis Div %LisDiv"
PushI O, "$Div Dept + Division"
PushI O, "$DivNm LongDies"
PushI O, "$DivSeq Seq"
PushI O, "$DivSts Status"
PushI O, "============================================ #Sto"
PushI O, "?sel Sto StoNm StoCNm"
PushI O, "fm   Location"
PushI O, "into #Sto"
PushI O, "?wh in strLis Loc %LisLoc"
PushI O, "$Sto"
PushI O, "$StoNm"
PushI O, "$StoCNm"
PushI O, "============================================= #Crd"
PushI O, "?sel        Crd CrdNm"
PushI O, "fm          Location"
PushI O, "into        #Crd"
PushI O, "?wh in nbrLis Crd %LisCrd"
PushI O, "$Crd"
PushI O, "$CrdNm"
PushI O, "============================================= #Oup"
PushI O, "sel  ?Crd ?CrdNm ?Mbr ?Age ?Sex ?Sts ?Dist ?Area ?Div ?DivNm ?Sto ?StoNm ?StoCNm ?Y ?M ?W ?WD ?D ?Dte Amt Qty Cnt"
PushI O, "into #Oup"
PushI O, "fm   #Tx x"
PushI O, "left #Crd a on x.Crd = a.Crd"
PushI O, "left #Div b on x.Div = b.Div"
PushI O, "left #Sto c on x.Sto = c.Sto"
PushI O, "left #MbrDta d on x.Mbr = d.Mbr"
PushI O, "wh   JCMCode in (Select Mbr From #TxMbr)"
PushI O, "============================================ #Cnt"
PushI O, "sel ?MbrCnt RecCnt TxCnt Qty Amt"
PushI O, "into #Cnt"
PushI O, "fm  #Tx"
PushI O, "$MbrCnt?Count(Distinct Mbr)"
PushI O, "$RecCnt Count(*)"
PushI O, "$TxCnt  Sum(TxCnt)"
PushI O, "$Qty    Sum(Qty)"
PushI O, "$Amt    Sum(Amt)"
PushI O, "============================================"
PushI O, "--"
PushI O, "============================================"
PushI O, "df eror fs--"
PushI O, "============================================"
PushI O, "-- EQ & NE t1 only TxtPm is allowed"
PushI O, "--         t2 allow TxtPm, *BLANK, and other text"
PushI O, "?LvlY    EQ %SumLvl Y"
PushI O, "?LvlM    EQ %SumLvl M"
PushI O, "?LvlW    EQ %SumLvl W"
PushI O, "?LvlD    EQ %SumLvl D"
PushI O, "?Y       OR ?LvlD ?LvlW ?LvlM ?LvlY`"
Samp_SqTp = JnCrLf(O)
End Property
