VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEmplID 
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Text            =   "5"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtTraining 
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Text            =   "ADS"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtSkill 
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Text            =   "2"
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtPhone 
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Text            =   "704-555-1234"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtTitle 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Text            =   "Boss"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtLastName 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Text            =   "End"
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtFirstName 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Bob"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtContact 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "Contact"
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddContact 
      Caption         =   "Add Contact"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   2160
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddContact_Click()
Dim rs As ADODB.Recordset
Set rs = CreateRS(1)

rs.AddNew
rs!FirstName = txtFirstName.Text
rs!LastName = txtLastName.Text
rs!Title = txtTitle.Text
rs!emplid = CInt(txtEmplID.Text)
rs!Phone = txtPhone.Text
rs!Skill = txtSkill.Text
rs!Training = txtTraining.Text
rs.Update

End Sub

Private Sub cmdEnd_Click()
Unload Me
End Sub
Private Function CreateRS(iIndex As Integer) As ADODB.Recordset
Dim rs As ADODB.Recordset
'========================================================================
'This function creates an ADO recordset used to pass contact/customer data to
'the DBServer.
'iIndex selects the proper rs.
'========================================================================

Set rs = New ADODB.Recordset

With rs.Fields
    Select Case iIndex
        Case 0 'Company
            .Append "ID", adBigInt
            .Append "sName", adChar, 200
            .Append "Address", adChar, 100
            .Append "City", adChar, 50
            .Append "State", adChar, 50
            .Append "Zip", adChar, 50
            .Append "Country", adChar, 50
            .Append "Phone", adChar, 100
            .Append "Fax", adChar, 50
            .Append "EmplID", adBigInt
            .Append "Type", adChar, 50
        Case 1 'Contact
            .Append "FirstName", adChar, 50
            .Append "LastName", adChar, 50
            .Append "Title", adChar, 50
            .Append "EmplID", adBigInt
            .Append "Phone", adChar, 100
            .Append "Skill", adChar, 50
            .Append "Training", adChar, 50
        Case 2 'Call Type
            .Append "CallCode", adChar, 50
            .Append "EmplID", adBigInt
        Case 3 'Call
            .Append "CustomerID", adBigInt
            .Append "ContactID", adBigInt
            .Append "CallCodeID", adBigInt
            .Append "ProductID", adBigInt
            .Append "EmplID", adBigInt
            .Append "NoteDate", adDBDate
            .Append "Note", adChar, 500
            .Append "EntryDate", adDBTimeStamp
            .Append "CallTime", adInteger
            .Append "CaseID", adBigInt
        Case 4 'Link Table
            .Append "CustomerID", adBigInt
            .Append "ContactID", adBigInt
            .Append "EmplID", adBigInt
    End Select
End With
rs.Open ' , , adOpenUnspecified

MsgBox rs.Fields.Count
Set CreateRS = rs
Set rs = Nothing

End Function

