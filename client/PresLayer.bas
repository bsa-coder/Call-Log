Attribute VB_Name = "PresLayer"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC89502B8"
Option Explicit

Const clCOMPANYBYDATE As Integer = 0
Const clCONTACT As Integer = 1
Const clPRODUCT As Integer = 2
Const clCALLCODE As Integer = 3
Const clCALL As Integer = 4
Const clHISTORY2 As Integer = 4
Const clLINK As Integer = 5
Const clHISTORY As Integer = 6
Const clCALLTYPE As Integer = 7
Const clEMPLOYEE As Integer = 8
Const clCOMPANYUPDATE As Integer = 9
Const clCONTACTUPDATE As Integer = 10
Const clCONTACTBYID As Integer = 12
Const clCOMPANYBYID As Integer = 13

Private Const TVM_GETCOUNT = &H1105&
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                        (ByVal hwnd As Long, _
                        ByVal wMsg As Long, ByVal _
                        wParam As Long, _
                        lParam As Any) As Long
Dim BLServer As CallTrackerBLServer.BLServer
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
Dim sTBText As String
Dim iCounter As Integer

'Create BL Server if none exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
End If

'Get all the calls for the company
'Set rsCalls = BLServer.GetHistory(CompanyID)
fMainForm.mfgHistory.AddItem "No Data", 1
For iCounter = 1 To fMainForm.mfgHistory.Rows
    If fMainForm.mfgHistory.Rows > 2 Then
        fMainForm.mfgHistory.RemoveItem 2
    Else
        Exit For
    End If
Next iCounter

If lCase <> 0 Then 'getting history by case number
    Set rsCalls = BLServer.GetLbData2(clHISTORY2, CStr(lCase))
Else 'getting history by company id
    Set rsCalls = BLServer.GetLbData2(clHISTORY, CStr(CompanyID))
End If
'sTBText = ""

'If Not ((rsCalls Is Nothing) Or (rsCalls.BOF And rsCalls.EOF)) Then
'    sTBText = "-----------------------------------------------" & vbCrLf
''    sTBText = sTBText & "Company Name: " & rsCalls!CompanyName & vbCrLf
'    sTBText = sTBText & "===============================================" & vbCrLf & vbCrLf
'
'    Do Until rsCalls.EOF
'        If lCase = 0 Then
'            sText = rsCalls!NoteDate & " (" & rsCalls!iCallTime & " min) " & rsCalls!ContactName & vbCrLf
'            sText = sText & "-----------------------------------------------" & vbCrLf
'            sText = sText & rsCalls!ProductName & "  :  "
'            sText = sText & rsCalls!CallType & " (" & rsCalls!LastName & ")" & " Case: " & rsCalls!CaseID & vbCrLf
'            sText = sText & "..............................................." & vbCrLf
'            sText = sText & rsCalls!sNote & vbCrLf
'            sText = sText & "===============================================" & vbCrLf
'        Else
'            If rsCalls!CaseID = lCase Then
'                sText = rsCalls!NoteDate & " (" & rsCalls!iCallTime & " min) " & rsCalls!ContactName & vbCrLf
'                sText = sText & "-----------------------------------------------" & vbCrLf
'                sText = sText & rsCalls!ProductName & "  :  "
'                sText = sText & rsCalls!CallType & " (" & rsCalls!LastName & ")" & " Case: " & rsCalls!CaseID & vbCrLf
'                sText = sText & "..............................................." & vbCrLf
'                sText = sText & rsCalls!sNote & vbCrLf
'                sText = sText & "===============================================" & vbCrLf
'            Else
'                sText = ""
'            End If
'        End If
'        sTBText = sTBText & sText
'        DoEvents
'        If Len(sTBText) > 32000 Then Exit Do
'        rsCalls.MoveNext
'    Loop
'Else
'    sTBText = ""
'End If

'tb.Text = sTBText
If ShowDetail(rsCalls) Then
End If
'ShowCallData (rsCalls)

End Sub
Sub ShowCallData(rs As ADODB.Recordset)
If Not ((rs Is Nothing) Or (rs.BOF And rs.EOF)) Then
'    txtEntry(0).Text = rsCalls!CompanyName
'    txtEntry(1).Text = rsCalls!ContactName
'    txtCallNote(0).Text = rsCalls!sNote
'    txtEnterCaseID.Text = rsCalls!CaseID
'    Do Until rsCalls.EOF
'        If lCase = 0 Then
'            sText = rsCalls!NoteDate & " (" & rsCalls!iCallTime & " min) " & rsCalls!ContactName & vbCrLf
'            sText = sText & "-----------------------------------------------" & vbCrLf
'            sText = sText & rsCalls!ProductName & "  :  "
'            sText = sText & rsCalls!CallType & " (" & rsCalls!LastName & ")" & " Case: " & rsCalls!CaseID & vbCrLf
'            sText = sText & "..............................................." & vbCrLf
'            sText = sText & "===============================================" & vbCrLf
'        Else
'            If rsCalls!CaseID = lCase Then
'                sText = rsCalls!NoteDate & " (" & rsCalls!iCallTime & " min) " & rsCalls!ContactName & vbCrLf
'                sText = sText & "-----------------------------------------------" & vbCrLf
'                sText = sText & rsCalls!ProductName & "  :  "
'                sText = sText & rsCalls!CallType & " (" & rsCalls!LastName & ")" & " Case: " & rsCalls!CaseID & vbCrLf
'                sText = sText & "..............................................." & vbCrLf
'                sText = sText & rsCalls!sNote & vbCrLf
'                sText = sText & "===============================================" & vbCrLf
'            Else
'                sText = ""
'            End If
'        End If
'        sTBText = sTBText & sText
'        DoEvents
'        If Len(sTBText) > 32000 Then Exit Do
'        rsCalls.MoveNext
'    Loop
'Else
'    sTBText = ""
End If
'
'tb.Text = sTBText

End Sub
Function ShowDetail(rs As ADODB.Recordset) As Boolean
Dim sGridString As String
Dim iCount As Integer
Dim iCounter As Integer
Dim iNumberOfFields As Integer

On Error GoTo ErrorHandler

If Not (rs.BOF Or rs.EOF) Then
rs.MoveFirst
ShowDetail = False
sGridString = ""
iNumberOfFields = rs.Fields.Count - 1
    
With fMainForm.mfgHistory
    .Redraw = False
    .Clear
    .Cols = 11 'rs.Fields.Count
    .Rows = 0
    .AddItem sGridString
    
'Load search detail
    For iCounter = 0 To 10 'iNumberOfFields
        .ColAlignment(iCounter) = 1
        .MergeCol(iCounter) = True     ' Allow merge on Columns 0 thru 3
    Next iCounter
    
    For iCount = 1 To rs.RecordCount
        
        sGridString = sGridString & rs!NoteDate & vbTab
        sGridString = sGridString & rs!sNote & vbTab
'        sGridString = sGridString & rs!CaseID & " (" & rs!recordid & ")" & vbTab
        sGridString = sGridString & rs!CaseID & vbTab
        sGridString = sGridString & rs!recordid & vbTab
        sGridString = sGridString & rs!LastName & vbTab
'        sGridString = sGridString & rs!ContactName & " (" & rs!ContactID & ")" & vbTab
        sGridString = sGridString & rs!ContactName & vbTab
        sGridString = sGridString & rs!ContactID & vbTab
        sGridString = sGridString & rs!ProductName & vbTab
'        sGridString = sGridString & rs!CallType & " (" & rs!CallCodeID & ")" & vbTab
        sGridString = sGridString & rs!CallType & vbTab
        sGridString = sGridString & rs!CallCodeID & vbTab
        sGridString = sGridString & rs!iCallTime & vbTab
        
        .AddItem sGridString
        .RowHeight(.Rows - 1) = 1000
        rs.MoveNext
        sGridString = ""
    Next iCount

    .FixedRows = 1
'Load heading row
    .TextMatrix(0, 0) = "Date"
    .ColWidth(0) = 1000
    .TextMatrix(0, 1) = "Note"
    .ColWidth(1) = 4000
    .TextMatrix(0, 2) = "CaseID"
    .ColWidth(2) = 500
    .TextMatrix(0, 3) = "RecordID"
    .ColWidth(3) = 0
    .TextMatrix(0, 4) = "PSG Engr"
    .ColWidth(4) = 1000
    .TextMatrix(0, 5) = "Contact"
    .ColWidth(5) = 1000
    .TextMatrix(0, 6) = "ContactID"
    .ColWidth(6) = 0
    .TextMatrix(0, 7) = "Product"
    .ColWidth(7) = 1000
    .TextMatrix(0, 8) = "Call Type"
    .ColWidth(8) = 1000
    .TextMatrix(0, 9) = "Call Code ID"
    .ColWidth(9) = 0
    .TextMatrix(0, 10) = "Duration"
    .ColWidth(10) = 500
        
    .Redraw = True
End With

ShowDetail = True
End If
Exit Function

ErrorHandler:
    MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::tvwCustomers"
End Function
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

Function ResizeControls() As Integer
Dim CmdLeftPosition As Integer
Dim CmdHeight As Integer
Dim CmdSpace As Integer
Dim StandardSpace As Integer
Dim FrameWidth As Integer
Dim FrameHeight As Integer

ResizeControls = 0
StandardSpace = 120

With fMainForm
    If fMainForm.WindowState <> 1 Then
        If .Width < 12240 Then .Width = 12240
        If .Height < 9120 Then .Height = 9120
        
        CmdLeftPosition = .Width - .cmdUpdate.Width - 2 * StandardSpace
        CmdHeight = .cmdUpdate.Height
        CmdSpace = (.Height - (660 + .txtEnterCaseID.Top + .txtEnterCaseID.Height + CmdHeight * 8 + 2 * StandardSpace)) / 9
        
'Position command buttons LEFT - Working OK
        .shBLStatus.Left = CmdLeftPosition + 1480
        .lblCaseID.Left = CmdLeftPosition
        .Label3.Left = CmdLeftPosition
        .txtEnterCaseID.Left = CmdLeftPosition
        .cmdUpdate.Left = CmdLeftPosition
        .cmdEditCall.Left = CmdLeftPosition
        .cmdNewCustomer.Left = CmdLeftPosition
        .cmdNewContact.Left = CmdLeftPosition
        .cmdEditContact.Left = CmdLeftPosition
        .cmdQuery.Left = CmdLeftPosition
        .cmdForceTVWLoad.Left = CmdLeftPosition
        .cmdCancel.Left = CmdLeftPosition
        
'Position command buttons TOP - Working OK
        .cmdUpdate.Top = .txtEnterCaseID.Top + .txtEnterCaseID.Height + StandardSpace
        .cmdEditCall.Top = .cmdUpdate.Top + .cmdUpdate.Height + CmdSpace
        .cmdNewCustomer.Top = .cmdEditCall.Top + .cmdEditCall.Height + CmdSpace + CmdSpace
        .cmdNewContact.Top = .cmdNewCustomer.Top + .cmdNewCustomer.Height + CmdSpace
        .cmdEditContact.Top = .cmdNewContact.Top + .cmdNewContact.Height + CmdSpace
        .cmdQuery.Top = .cmdEditContact.Top + .cmdEditContact.Height + CmdSpace
        .cmdForceTVWLoad.Top = .cmdQuery.Top + .cmdQuery.Height + CmdSpace
        .cmdCancel.Top = .cmdForceTVWLoad.Top + .cmdForceTVWLoad.Height + CmdSpace + CmdSpace
        
'Size the FRAMES
        .frmCustomer.Width = CmdLeftPosition - 2 * StandardSpace
        .frmCall.Width = CmdLeftPosition - 2 * StandardSpace
        FrameWidth = .frmCustomer.Width - 3 * StandardSpace
        FrameHeight = 660 + .cmdTimer.Top + .cmdTimer.Height + 1 * StandardSpace
        FrameHeight = (.Height - FrameHeight)
        .frmCall.Height = FrameHeight * 0.45
        .frmCustomer.Height = FrameHeight * 0.55
        
        .lstItem(clPRODUCT).Height = .frmCall.Height - .lstItem(clPRODUCT).Top - StandardSpace
        .txtCallNote(0).Height = (.frmCall.Height - 3 * StandardSpace - .txtCallNote(0).Top) / 2
        .txtCallNote(1).Top = .txtCallNote(0).Top + .txtCallNote(0).Height + 2 * StandardSpace
        .txtCallNote(1).Height = .txtCallNote(0).Height
        .lblCallNote(1).Top = .txtCallNote(1).Top - 180
        
        .frmCustomer.Top = .cmdTimer.Top + .cmdTimer.Height + .frmCall.Height + 0 * StandardSpace
        
        .txtEntry(0).Width = FrameWidth * 0.4
        .txtEntry(clCONTACT).Width = .txtEntry(0).Width
        
        .txtCallNote(0).Left = .txtEntry(0).Width + StandardSpace + StandardSpace
        .txtCallNote(1).Left = .txtCallNote(0).Left
        .txtCallNote(0).Width = FrameWidth * 0.6
        .txtCallNote(1).Width = .txtCallNote(0).Width
        .lblCallNote(0).Left = .txtCallNote(0).Left
        .lblCallNote(1).Left = .txtCallNote(0).Left
        
        .tvwCustomers.Width = .txtEntry(0).Width
        .tvwCustomers.Height = (.frmCustomer.Height - .tvwCustomers.Top - 2 * StandardSpace) * 0.6
        .lstItem(1).Width = .txtEntry(0).Width
        .mfgHistory.Left = .txtCallNote(0).Left
        .mfgHistory.Width = .txtCallNote(0).Width
        .Label1.Left = .txtCallNote(0).Left
        .mfgHistory.Height = .frmCustomer.Height - .mfgHistory.Top - StandardSpace
        .lstItem(1).Top = .tvwCustomers.Top + .tvwCustomers.Height + StandardSpace / 2
        .lstItem(1).Height = .frmCustomer.Height - .lstItem(1).Top - StandardSpace
        
        .sbMain.Panels(3).Width = .Width - 3 * StandardSpace - (.sbMain.Panels(1).Width + .sbMain.Panels(2).Width + .sbMain.Panels(4).Width)
        
    End If
End With
ResizeControls = 1
End Function

