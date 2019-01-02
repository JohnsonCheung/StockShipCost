Attribute VB_Name = "MTp_Sq__Types"
Option Compare Database

Public Type PmRslt
    Pm As New Dictionary
    Er() As String
End Type


Public Type SqyRslt
    Sqy() As String
    Er() As String
End Type
Public Type SqlRslt
    Sql As String
    Er() As String
End Type

Public Type LyRslt
    Ly() As String
    Er() As String
End Type
Public Type LnxAyRslt
    LnxAy() As Lnx
    Er() As String
End Type
