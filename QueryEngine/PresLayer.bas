Attribute VB_Name = "PresLayer"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC89502B8"
Option Explicit
'##ModelId=3A0F61E9021A

Const clCOMPANY As Integer = 0
Const clCONTACT As Integer = 1
Const clPRODUCT As Integer = 2
Const clCALLCODE As Integer = 3
Const clCALL As Integer = 4
Const clLINK As Integer = 5
Const clHISTORY As Integer = 6
Const clCALLTYPE As Integer = 7
Const clEMPLOYEE As Integer = 8
Const clCOMPANYUPDATE As Integer = 9
Const clCONTACTUPDATE As Integer = 10
Const clCONTACT2 As Integer = 12
Const clCOMPANY2 As Integer = 13

Private Const TVM_GETCOUNT = &H1105&
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                        (ByVal hwnd As Long, _
                        ByVal wMsg As Long, ByVal _
                        wParam As Long, _
                        lParam As Any) As Long
Dim BLServer As CallTrackerBLServer.BLServer
'##ModelId=3A0F61E9022C
Sub GoToItem(lb As ListBox, ID As Long)
Dim iCounter As Integer
If lb.ListCount > 0 Then
    For iCounter = 0 To lb.ListCount - 1
        If lb.ItemData(iCounter) = ID Then
            lb.ListIndex = iCounter
            Exit For
        End If
    Next iCounter
End If
End Sub
'##ModelId=3A0F61E90345
Sub FindListItem(lb As ListBox, txt As TextBox)
Dim iCount As Integer
Dim sTemp As String
Dim str As String
Dim iInit As Integer

On Error GoTo ErrorHandler

sTemp = txt.Text

If Len(sTemp) = 0 Then Exit Sub

If Len(sTemp) > 1 Then iInit = lb.ListIndex Else iInit = 0

For iCount = iInit To lb.ListCount - 1
    str = lb.List(iCount)
    If UCase(Left(str, Len(sTemp))) = UCase(sTemp) Then
        lb.ListIndex = iCount
        Exit For
    End If
Next iCount

txt.Text = sTemp
txt.SelStart = Len(sTemp)
Exit Sub

ErrorHandler:
MsgBox "Error finding customer in lstCustomer", , "FProdSupLoading"
Resume Next

End Sub
'-----------------------------------------------------------------
'get the call history
'##ModelId=3A0F61EA006B
Sub GetCompanyHistory(CompanyID As Long, tb As TextBox, Optional lCase As Long)
Dim rsCalls As ADODB.Recordset
Dim sText As String

'Create BL Server if none exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
End If

'Get all the calls for the company
Set rsCalls = BLServer.GetHistory(CompanyID)

tb.Text = ""

If Not (rsCalls Is Nothing) Then
    Do Until rsCalls.EOF
        If lCase = 0 Then
            sText = rsCalls!NoteDate & " (" & rsCalls!iCallTime & " min) " & rsCalls!ContactName & vbCrLf
            sText = sText & "-----------------------------------------------" & vbCrLf
            sText = sText & rsCalls!ProductName & "  :  "
            sText = sText & rsCalls!CallType & " (" & rsCalls!LastName & ")" & " Case: " & rsCalls!CaseID & vbCrLf
            sText = sText & "..............................................." & vbCrLf
            sText = sText & rsCalls!sNote & vbCrLf
            sText = sText & "===============================================" & vbCrLf
        Else
            If rsCalls!CaseID = lCase Then
                sText = rsCalls!NoteDate & " (" & rsCalls!iCallTime & " min) " & rsCalls!ContactName & vbCrLf
                sText = sText & "-----------------------------------------------" & vbCrLf
                sText = sText & rsCalls!ProductName & "  :  "
                sText = sText & rsCalls!CallType & " (" & rsCalls!LastName & ")" & " Case: " & rsCalls!CaseID & vbCrLf
                sText = sText & "..............................................." & vbCrLf
                sText = sText & rsCalls!sNote & vbCrLf
                sText = sText & "===============================================" & vbCrLf
            Else
                sText = ""
            End If
        End If
        tb.Text = tb.Text & sText
        
        rsCalls.MoveNext
        DoEvents
    Loop
Else
    tb.Text = ""
End If
    
End Sub
'##ModelId=3A0F61EA0184
Function LoadListBox(lb As ListBox, LBs As Object, Optional Filter As Integer) As Integer
Dim ID As Integer
Dim vCounter As Variant
On Error Resume Next

'Initialize listbox
LoadListBox = 0
If (Filter < 1) Then
    lb.Clear
    For Each vCounter In LBs
        lb.AddItem vCounter.sName
        lb.ItemData(lb.NewIndex) = vCounter.ID
    Next vCounter
Else
    For Each vCounter In LBs
        If vCounter.ID = Filter Then
            lb.AddItem vCounter.sName
            lb.ItemData(lb.NewIndex) = vCounter.ID
        End If
    Next vCounter
End If
lb.Refresh
LoadListBox = 1
End Function
'##ModelId=3A0F61EA031E
Function LoadComboBox(cboBox As ComboBox, oData As Object) As Integer
Dim vCounter As Variant
Dim sTemp As String

On Error Resume Next

'Initialize listbox
cboBox.Clear
LoadComboBox = 0

For Each vCounter In oData
    cboBox.AddItem vCounter.sName
    cboBox.ItemData(cboBox.NewIndex) = vCounter.ID
Next vCounter

cboBox.AddItem "Select Employee", 0
cboBox.ItemData(cboBox.NewIndex) = 0
cboBox.ListIndex = 0

LoadComboBox = 1
End Function


