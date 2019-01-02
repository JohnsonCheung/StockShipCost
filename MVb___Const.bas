Attribute VB_Name = "MVb___Const"
Option Compare Binary
Option Explicit
Public Const vbOpnBkt$ = "("
Public Const vbDblQuote$ = """"
Public Const vbOpnSqBkt$ = "["
Public Fso As New Scripting.FileSystemObject
Const H$ = "C:\Users\User\Desktop\SAPAccessReports\"
Const H1$ = "C:\Users\User\Desktop\"
'------------------------------------------
'From:
'https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/sql/sql-server-express-user-instances
Public Const Samp_CnStr_SQLEXPR$ = "Data Source=.\\SQLExpress;Integrated Security=true;" & _
"User Instance=true;AttachDBFilename=|DataDirectory|\InstanceDB.mdf;" & _
"Initial Catalog=InstanceDB;"
'------------------------------------------
'From:
'https://social.msdn.microsoft.com/Forums/vstudio/en-US/61d45bef-eea7-4366-a8ad-e15a1fa3d544/vb6-to-connect-with-sqlexpress?forum=vbgeneral
Public Const Samp_CnStr_SQLEXPR_NotWrk3$ = _
"Provider=SQLNCLI.1;Integrated Security=SSPI;AttachDBFileName=C:\User\Users\northwnd.mdf;Data Source=.\sqlexpress"
Public Const Samp_CnStr_ADO_SAMP_SQL_EXPR_NOT_WRK$ = _
"Provider=Sq_LoleDb;Integrated Security=SSPI;AttachDBFileName=C:\User\Users\northwnd.mdf;Data Source=.\sqlexpress"
'--------------------------------
'From https://social.msdn.microsoft.com/Forums/en-US/a73a838b-ec3f-419b-be65-8b1732fbf4d0/connect-to-a-remote-sql-server-db?forum=isvvba
Public Const Samp_CnStr_SQLEXPR_NotWrk1$ = "driver={SQL Server};" & _
      "server=LAPTOP-SH6AEQSO;uid=MyUserName;pwd=;database=pubs"
   
Public Const Samp_CnStr_SQLEXPR_NotWrk2$ = "driver={SQL Server};" & _
      "server=127.0.0.1;uid=MyUserName;pwd=;database=pubs"
   
Public Const Samp_CnStr_SQLEXPR_NotWrk$ = ".\SQLExpress;AttachDbFilename=c:\mydbfile.mdf;Database=dbname;" & _
"Trusted_Connection=Yes;"
'"Typical normal SQL Server connection string: Data Source=myServerAddress;
'"Initial Catalog=myDataBase;Integrated Security=SSPI;"

'From VisualStudio
Public Const SampSqlCnStr_NotWrk$ = _
    "Data Source=LAPTOP-SH6AEQSO\ProjectsV13;Initial Catalog=master;Integrated Security=True;Connect Timeout=30;" & _
    "Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False"

Public Const Samp_Fb_Duty_Dta$ _
                                    = H & "DutyPrepay5\DutyPrepay5_Data.mdb"
Public Const Samp_Fb_Duty_Pgm$ _
                                    = H & "DutyPrepay5\DutyPrepay5.accdb"
Public Const Samp_Fx_KE24 _
                                    = H & "DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls"
Public Const Samp_Fb_Duty_PgmBackup$ _
                                    = H & "DutyPrepay5\DutyPrepay5_Backup.accdb"
Public Const Samp_Fb_TaxCmp$ _
                                    = H1 & "QFinalSln\TaxExpCmp v1.3.accdb"
Public Const Samp_Fb_ShpRate$ _
                                    = H1 & "QFinalSln\StockShipRate (ver 1.0).accdb"
Property Get DbEng() As DBEngine
Set DbEng = DAO.DBEngine
End Property
Private Function Fb_Db(A) As DAO.Database
Set Fb_Db = DAO.DBEngine.OpenDatabase(A)
End Function
Property Get Samp_Cn_Duty_Dta() As ADODB.Connection
Set Samp_Cn_Duty_Dta = Fb_Cn(Samp_Fb_Duty_Dta)
End Property

Property Get Samp_Db_Duty_Dta() As Database
Static Y As Database
If IsNothing(Y) Then Set Y = Fb_Db(Samp_Fb_Duty_Dta)
Set Samp_Db_Duty_Dta = Y
End Property

Sub XX()
Dim A
'{00024500-0000-0000-C000-000000000046}
Set A = Interaction.CreateObject("{00024500-0000-0000-C000-000000000046}", "Excel.Application")
Stop
End Sub

