Attribute VB_Name = "MIde_Z_Md_Op_Mov"
Option Compare Binary
Option Explicit
Sub Md_XMov(A As CodeModule)
'Move the MdNm in SrcPj-(Lib_XX) to TarPj-(VbLib)
'Ass ZMd_NotExist_InTar
Dim SrcCmp As VBComponent
Dim TmpFil$
    TmpFil = TmpFfn(".txt")
'    Set SrcCmp = ZSrcCmp
    SrcCmp.Export TmpFil
    If SrcCmp.Type = vbext_ct_ClassModule Then
        'ZRmvFst4Lines TmpFil
    End If
Dim TarCmp As VBComponent
'    Set TarCmp = ZTarPj.VBComponents.Add(ZMd_Ty)
    TarCmp.CodeModule.AddFromFile TmpFil
'ZSrcPj.VBComponents.Remove SrcCmp
Kill TmpFil
End Sub

Sub MdLikNm_XMov(MdLikNm$)
Dim I
For Each I In AyNz(CurPj_MdAy(WhMd(Nm:=WhNm("^" & MdLikNm))))
    Md_XMov CvMd(I)
Next
End Sub
