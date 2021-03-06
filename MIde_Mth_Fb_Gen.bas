Attribute VB_Name = "MIde_Mth_Fb_Gen"
Option Compare Binary
Option Explicit

Sub GenCrtDistMth()
WSetDb MthDb
WDrp "DistMth #A #B"
Q = "Select Distinct Nm,Count(*) as LinesIdCnt Into DistMth from DistLines group by Nm": W.Execute Q
Q = "Alter Table DistMth Add Column LinesIdLis Text(255), LinesLis Memo, ToMd Text(50)": W.Execute Q
Wt_XCrt_FldLisTbl "DistLines", "#A", "Nm", "LinesId", " ", True
Wt_XCrt_FldLisTbl "DistLines", "#B", "Nm", "Lines", vbCrLf & vbCrLf, True
Q = "Update DistMth x inner join [#A] a on x.Nm = a.Nm set x.LinesIdLis = a.LinesIdLis":                W.Execute Q
Q = "Update DistMth x inner join [#B] a on x.Nm = a.Nm set x.LinesLis = a.LinesLis":                    W.Execute Q
Q = "Update DistMth x inner join MthLoc a on x.Nm = a.Nm set x.ToMd = IIf(a.ToMd='','AAMod',a.ToMd)":   W.Execute Q
End Sub

Sub GenCrtMdDic()
WSetDb MthDb
WDrp "MdDic"
Wt_XCrt_FldLisTbl "DistMth", "MdDic", "ToMd", "LinesLis", vbCrLf & vbCrLf, True, "Lines"
End Sub

Sub UpdMthLoc()
WSetDb MthDb
WDrp "#A"
Q = "Select x.Nm into [#A] from DistMth x left join MthLoc a on x.Nm=a.Nm where IsNull(a.Nm)": W.Execute Q
Q = "Insert into MthLoc (Nm) Select Nm from [#A]": W.Execute Q
End Sub

Private Sub ZZ()
GenCrtDistMth
GenCrtMdDic
UpdMthLoc
End Sub

Private Sub Z()
End Sub
