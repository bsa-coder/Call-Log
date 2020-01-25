VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FNewMain 
   Caption         =   "New Call Entry Form"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   6225
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TreeView tvwCurrent 
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   7435
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   353
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imgTVWPictures"
      Appearance      =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgdCallHistory 
      Height          =   4215
      Left            =   3600
      TabIndex        =   3
      Top             =   600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7435
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtCompany 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      TabIndex        =   2
      Text            =   "Company"
      Top             =   120
      Width           =   3855
   End
   Begin VB.TextBox txtContact 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   1
      Text            =   "Contact"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   420
      Left            =   7680
      TabIndex        =   0
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Shape shBLStatus 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9240
      Shape           =   3  'Circle
      Top             =   120
      Width           =   135
   End
   Begin ComctlLib.ImageList imgTVWPictures 
      Left            =   0
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":094E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":0C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":0F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":129C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":15B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":18D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":1BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":1F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":221E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":2538
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":2852
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":2B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNewMain.frx":2E86
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FNewMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsGeneric As ADODB.Recordset

Dim WithEvents BLServer As CallTrackerBLServer.BLServer
Attribute BLServer.VB_VarHelpID = -1

Private Sub txtCompany_Change()
    If Len(txtCompany.Text) > 3 Then
        'GetListOfCompanies
    End If
End Sub

Private Sub txtCompany_GotFocus()
    With txtCompany
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCompany_LostFocus()
    'GetListOfCompanies
End Sub

Private Sub txtContact_Change()
    If Len(txtContact.Text) > 3 Then
        'GetListOfContacts
    End If
End Sub

Private Sub txtContact_GotFocus()
    With txtContact
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtContact_LostFocus()
    'GetListOfContacts
End Sub

Private Function GetListOfContacts() As Recordset
Dim sSelect As String
Dim sFrom As String
Dim sWhere As String
Dim sOrderBy As String
Dim sEnd As String
Dim Criteria As String
Dim rsCompanies As ADODB.Recordset
Dim iCounter As Integer
Dim iIndex As Integer

'Get list of companies
sSelect = "SELECT Company.ID, First([Company]![CompanyName] & ', ' & [Company]![Address]"
sSelect = sSelect & "', ' & [Company]![City] & ', ' & [Company]![State]) AS CompName "

sFrom = "FROM Contact RIGHT JOIN ((Company INNER JOIN "
sFrom = sFrom & "[Contact Link] ON Company.ID = [Contact Link].CompanyID) "
sFrom = sFrom & "INNER JOIN SupportCalls ON Company.ID = SupportCalls.CompanyID) ON "
sFrom = sFrom & "Contact.ID = [Contact Link].ContactID "

sWhere = "WHERE (((Contact.FirstName) Like '*') And "
sWhere = sWhere & "((SupportCalls.NoteDate) > Now() - 160)) "

sOrderBy = "GROUP BY Company.ID"

sEnd = ";"

Criteria = sSelect & sFrom & sWhere & sOrderBy & sEnd

frmSplash.MousePointer = vbHourglass
iCounter = 0

'Make sure a server exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    shBLStatus.FillColor = vbGreen
    If BLServer Is Nothing Then
        shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server" & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::GetListOfContacts"
        Exit Function
    End If
End If

'Get the customer data from the server
Set rsGeneric = BLServer.GetLbData(iIndex)
If rsGeneric Is Nothing Then
    MsgBox "Recordset not created"
    Exit Function
End If
Set BLServer = Nothing

End Function
