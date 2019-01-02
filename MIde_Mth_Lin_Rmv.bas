Attribute VB_Name = "MIde_Mth_Lin_Rmv"
Option Compare Binary
Option Explicit
Function XRmv_Mdy$(A)
XRmv_Mdy = LTrim(XRmv_PfxAySpc(A, MdyAy))
End Function

Function RmvMthTy$(A)
RmvMthTy = XRmv_PfxAySpc(A, MthTyAy)
End Function
