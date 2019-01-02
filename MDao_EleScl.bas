Attribute VB_Name = "MDao_EleScl"
Option Compare Binary
Option Explicit

Function TF_EleScl$(T, F)
TF_EleScl = Dbtf_EleScl(CurDb, T, F)
End Function

