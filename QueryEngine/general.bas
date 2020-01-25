Attribute VB_Name = "MGeneral"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC8A001F6"
Option Explicit

'##ModelId=39EBC8A0034C
Public fMainForm As FMain

'##ModelId=39EBC8A003AE
Sub Main()
Dim dtStartTime As Date
Dim dtDisplayTime As Variant

    dtDisplayTime = 2 / 3600
    dtDisplayTime = dtDisplayTime / 24

    dtStartTime = Now()

'=======================================================
' Add this code back in when restricted access is required.
'    Dim fLogin As New frmLogin
'    fLogin.Show vbModal
'    If Not fLogin.OK Then
        'Login Failed so exit app
'        End
'    End If
'    Unload fLogin
'=======================================================

    frmSplash.Show
    frmSplash.Refresh
    
' Check for DB location
    
    Set fMainForm = New FMain
    
'    Do Until Now() >= dtStartTime + dtDisplayTime
'        frmSplash.Label1.Caption = CStr(Format(Now(), "HH:MM:SS"))
'        frmSplash.Refresh
'    Loop
    
    Load fMainForm
    
    Unload frmSplash

    fMainForm.Show
End Sub

