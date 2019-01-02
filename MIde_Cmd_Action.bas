Attribute VB_Name = "MIde_Cmd_Action"
Option Compare Binary
Option Explicit

Sub XTile_H()
WinTileVBtn.Execute
End Sub

Sub XTile_V()
WinTileVBtn.Execute
End Sub

Property Get TileVBtn() As CommandBarButton
Dim O As CommandBarButton
Set O = WinPop.CommandBar.Controls(3)
If O.Caption <> "Tile &Vertically" Then Stop
Set TileVBtn = O
End Property
