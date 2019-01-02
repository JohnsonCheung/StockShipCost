Attribute VB_Name = "MIde_Mth_Ix_FT"
Option Compare Binary
Option Explicit
Function Src_MthFTIxAyIxAy(A$()) As FTIx()
Dim Ix
For Each Ix In AyNz(Src_MthIxAy(A))
    PushObj Src_MthFTIxAyIxAy, New_FTIx(Ix, SrcMthIx_MthIxTo(A, Ix))
Next
End Function

Function SrcMthNm_FmCnt_FstIxAy(A$(), MthNm) As FTIx()
Dim IxAy&(), F&, T&
IxAy = SrcMthNm_MthIxAy(A, MthNm)
Dim J%
For J = 0 To UB(IxAy)
   F = IxAy(J)
   T = SrcMthFmIx_MthToIx(A, F)
   Push SrcMthNm_FmCnt_FstIxAy, New_FTIx(F, T)
Next
End Function
