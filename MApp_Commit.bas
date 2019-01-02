Attribute VB_Name = "MApp_Commit"
Option Compare Binary
Option Explicit

Sub XCommit(Optional Msg$ = "Commit")
'Pj_XExp CurPj
WXEns_Cxt_COMMIT
Fcmd_XRun_Max WFcmd_COMMIT, Msg
End Sub

Property Get WFcmd_COMMIT$()
WFcmd_COMMIT = WPth & "Commit.Cmd"
End Property

Property Get WFcmd_PUSH$()
WFcmd_PUSH = WPth & "Pushing.Cmd"
End Property

Sub XPush_App()
WXEns_Cxt_PUSH
Fcmd_XRun_Max WCxt_PUSH
End Sub

Private Sub WXEns_Cxt_COMMIT()
Str_XWrt WCxt_COMMIT, WFcmd_COMMIT
End Sub

Private Sub WXEns_Cxt_PUSH()
Str_XWrt WCxt_PUSH, WFcmd_PUSH
End Sub

Private Property Get WCxt_COMMIT$()
Dim O$(), Cd$, GitAdd$, GitCommit$, GitPush, T$
Cd = QQ_Fmt("Cd ""?""", SrcPth)
GitAdd = "git add -A"
GitCommit = "git commit --message=%1%"
Push O, Cd
Push O, GitAdd
Push O, GitCommit
Push O, "Pause"
WCxt_COMMIT = JnCrLf(O)
End Property

Private Property Get WCxt_PUSH$()
Dim O$(), Cd$, GitPush, T
Cd = QQ_Fmt("Cd ""?""", SrcPth)
GitPush = "git push -u https://johnsoncheung@github.com/johnsoncheung/StockShipRate.git master"
Push O, Cd
Push O, GitPush
Push O, "Pause"
WCxt_PUSH = JnCrLf(O)
End Property

Sub XExp()
CurVbe_XExp
End Sub

Sub XBrw_Fcmd_COMMIT()
WXEns_Cxt_COMMIT
Ft_XBrw WFcmd_COMMIT
End Sub

Sub XBrw_Fcmd_PUSH()
WXEns_Cxt_PUSH
Ft_XBrw WFcmd_PUSH
End Sub

Private Sub WX()
Dim O$()
Push O, "echo ""# QLib"" >> README.md"
Push O, "git Init"
Push O, "git add README.md"
Push O, "git commit -m ""first commit"""
Push O, "git remote add origin https://github.com/JohnsonCheung/QLib.git"
Push O, "git push -u origin master"

End Sub

Private Sub WX1()
'git remote add origin https://github.com/JohnsonCheung/QLib.git
'git push -u origin master
End Sub

Private Sub ZZ()
Dim A$
XCommit A
XPush_App
XExp
End Sub
