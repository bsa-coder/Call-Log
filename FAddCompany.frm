VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FAddItem 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Add Company, Contact, or Code"
   ClientHeight    =   6450
   ClientLeft      =   2820
   ClientTop       =   2580
   ClientWidth     =   10050
   Icon            =   "FAddCompany.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   10050
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   3840
      Width           =   3735
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   3735
   End
   Begin VB.Frame frmContact 
      Height          =   3615
      Left            =   240
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   7695
      Begin VB.Frame Frame3 
         Caption         =   "Training"
         Height          =   1695
         Left            =   4920
         TabIndex        =   55
         Top             =   1800
         Width           =   2655
         Begin VB.CheckBox chkTraining 
            Caption         =   "Comms (adv)"
            Height          =   255
            Index           =   6
            Left            =   1200
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CheckBox chkTraining 
            Caption         =   "Link Dev."
            Height          =   255
            Index           =   5
            Left            =   1200
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   960
            Width           =   1095
         End
         Begin VB.CheckBox chkTraining 
            Caption         =   "Runtime"
            Height          =   255
            Index           =   4
            Left            =   1200
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chkTraining 
            Caption         =   "Servo"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox chkTraining 
            Caption         =   "DC"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   960
            Width           =   975
         End
         Begin VB.CheckBox chkTraining 
            Caption         =   "AC"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkTraining 
            Caption         =   "None"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   360
            Value           =   1  'Checked
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Skill Level"
         Height          =   1695
         Left            =   3360
         TabIndex        =   51
         Top             =   1800
         Width           =   1335
         Begin VB.OptionButton optSkill 
            Caption         =   "Skilled"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1095
         End
         Begin VB.OptionButton optSkill 
            Caption         =   "Medium"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton optSkill 
            Caption         =   "Novice"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Contact Numbers"
         Height          =   1695
         Left            =   120
         TabIndex        =   38
         Top             =   1800
         Width           =   3015
         Begin VB.OptionButton optPrimary 
            Height          =   255
            Index           =   3
            Left            =   2640
            TabIndex        =   50
            TabStop         =   0   'False
            ToolTipText     =   "Select PAGER as primary number"
            Top             =   1320
            Width           =   255
         End
         Begin VB.OptionButton optPrimary 
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   49
            TabStop         =   0   'False
            ToolTipText     =   "Select CELL as primary number"
            Top             =   960
            Width           =   255
         End
         Begin VB.OptionButton optPrimary 
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "Select FAX as primary number"
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optPrimary 
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   47
            TabStop         =   0   'False
            ToolTipText     =   "Select PHONE as primary number"
            Top             =   240
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.TextBox txtContactNumber 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtContactNumber 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtContactNumber 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtContactNumber 
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   39
            Text            =   "999-999-9999"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblContactNumber 
            Alignment       =   1  'Right Justify
            Caption         =   "Pager: "
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   46
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblContactNumber 
            Alignment       =   1  'Right Justify
            Caption         =   "Cell: "
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblContactNumber 
            Alignment       =   1  'Right Justify
            Caption         =   "Fax: "
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   44
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblContactNumber 
            Alignment       =   1  'Right Justify
            Caption         =   "Phone: "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Text            =   "Enter the new contact information.  The items with a * are required."
         Top             =   240
         Width           =   7335
      End
      Begin VB.TextBox txtLastName 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1080
         Width           =   6015
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1440
         Width           =   6015
      End
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   720
         Width           =   6015
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "* First Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   765
         Width           =   1125
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Title"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1485
         Width           =   1125
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Name"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1125
         Width           =   1125
      End
   End
   Begin ComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   6195
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "7/17/2002"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "1:08 AM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   7752
            MinWidth        =   7762
            Text            =   "FAddItem"
            TextSave        =   "FAddItem"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmProduct 
      Height          =   3615
      Left            =   600
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox txtProductName 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   720
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Text            =   "Enter the new product name.  The items with a * are required."
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "* Product Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   765
         Width           =   1365
      End
   End
   Begin VB.Frame frmCallType 
      Height          =   3615
      Left            =   480
      TabIndex        =   35
      Top             =   480
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox txtCallType 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   720
         Width           =   5895
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Text            =   "Enter the new call type.  The items with a * are required."
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "* Call Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   765
         Width           =   1125
      End
   End
   Begin VB.Frame frmCompany 
      Height          =   3615
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   7695
      Begin VB.Frame Frame4 
         Caption         =   "Company Type (Select One)"
         Height          =   1455
         Left            =   3720
         TabIndex        =   63
         Top             =   2040
         Width           =   3855
         Begin VB.OptionButton optType 
            Caption         =   "Invensys Family"
            Height          =   255
            Index           =   7
            Left            =   1800
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1815
         End
         Begin VB.OptionButton optType 
            Caption         =   "Distributor"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton optType 
            Caption         =   "Rep"
            Height          =   255
            Index           =   5
            Left            =   1800
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton optType 
            Caption         =   "Eurotherm Family"
            Height          =   255
            Index           =   6
            Left            =   1800
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   840
            Width           =   1815
         End
         Begin VB.OptionButton optType 
            Caption         =   "Service"
            Height          =   255
            Index           =   4
            Left            =   1800
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optType 
            Caption         =   "Integrator"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton optType 
            Caption         =   "OEM"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton optType 
            Caption         =   "End User"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox txtCompanyName 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   720
         Width           =   6255
      End
      Begin VB.TextBox txtCountry 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtState 
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtZip 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtFax 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox txtCity 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1080
         Width           =   6255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Text            =   "Enter the new company information.  The items with a * are required."
         Top             =   240
         Width           =   7335
      End
      Begin VB.ComboBox cboState 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "FAddCompany.frx":0442
         Left            =   1200
         List            =   "FAddCompany.frx":0444
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "FAX"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   3285
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "* Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2925
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Country"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2570
         Width           =   1005
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1130
         Width           =   1005
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "ZIP"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2210
         Width           =   1005
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "State"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1850
         Width           =   1005
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "City"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1490
         Width           =   1005
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "* Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   770
         Width           =   1005
      End
   End
End
Attribute VB_Name = "FAddItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC89603D2"
Option Explicit
'##ModelId=39EBC8970080
Dim miEmpl As Long
Dim msSkill As String
Dim msTraining As String
Dim msPhone As String
Dim mbNewCompany As Boolean

'##ModelId=39EBC89700DA
Dim miCompany As Long
'##ModelId=39EBC8970134
Dim miContact As Long
'##ModelId=39EBC8970199
Dim miLink As Long
'##ModelId=39EBC8970207
Dim miCode As Long
'##ModelId=39EBC8970275
Dim miAction As Integer
'##ModelId=39EBC89702ED
Dim miError As Integer
'##ModelId=39EBC897036F
Dim msText As String
Dim mcType As String
'##ModelId=39EBC89800D2
Dim iUpdate As Integer

'##ModelId=39EBC8980013
Dim NewCompany As CCustomer
'##ModelId=39EBC898005C
Dim NewContact As CContact
'##ModelId=39EBC89800A0
Dim NewCallCode As CCallCode

'##ModelId=3A0F61EB032A
Dim WithEvents BLServer As CallTrackerBLServer.BLServer
Attribute BLServer.VB_VarHelpID = -1

Const clCOMPANYBYDATE As Integer = 0
Const clCONTACT As Integer = 1
Const clLINK As Integer = 2
Const clCALLTYPE As Integer = 3
Const clCALL As Integer = 4
Const clPRODUCT As Integer = 5
Const clHISTORY As Integer = 6
Const clCALLCODE As Integer = 7
Const clEMPLOYEE As Integer = 8
Const clCOMPANYUPDATE As Integer = 9
Const clCONTACTUPDATE As Integer = 10

Const clNEWCALLS As Integer = 17
Const clNEWCOMPANIES As Integer = 18
Const clNEWCONTACTS As Integer = 19
Const clNEWCOMPANIESNOCALLS As Integer = 20
Dim aState(50, 2) As String

'##ModelId=3A0F61EB033E
Private Sub AddNewItems()
Dim iTemp As Long

miError = -1

Select Case iUpdate
    Case clCOMPANYBYDATE + 1
    '===== Add the company ====================
        With NewCompany
            miCompany = BLServer.AddCustomer(.sName, .sAddress, .sCity, .sState, .sZIP, .sCountry, .sPhone, .sFAX, miEmpl, mcType)
        End With
        If miCompany = 0 Then
            miError = 0
            Exit Sub
        End If
    '===== Add the contact ====================
        With NewContact
            miContact = BLServer.AddContact(.sFirstName, .sLastName, .sTitle, miEmpl, .sPhone, .sSkill, .sTraining)
        End With
    '===== Add the Company/Contact Link =======
        If miContact = 0 Then
            miError = 0
            Exit Sub
        Else
            miLink = BLServer.AddCompanyLink(miCompany, miContact, miEmpl)
        End If
        If miLink = 0 Then
            miError = 0
            Exit Sub
        End If
    Case clCONTACT + 1
    '===== Add the contact ====================
        With NewContact
            miContact = BLServer.AddContact(.sFirstName, .sLastName, .sTitle, miEmpl, .sPhone, .sSkill, .sTraining)
        End With
        
    '===== Add the Company/Contact Link =======
        If ((miContact = 0) Or (miCompany = 0)) Then
            miError = 0
            Exit Sub
        Else
            miLink = BLServer.AddCompanyLink(miCompany, miContact, miEmpl)
        End If
        
    '===== Test if adding link worked =========
        If miLink = 0 Then
            miError = 0
            Exit Sub
        End If
    Case clCALLTYPE + 1
        miCode = BLServer.AddCallCode(txtCallType.Text, miEmpl)
        If miCode = 0 Then
            miError = 0
            Exit Sub
        End If
    Case clCOMPANYUPDATE + 1
    '===== Add the company ====================
        With NewCompany
            miCompany = BLServer.AddCustomer(.sName, .sAddress, .sCity, .sState, .sZIP, .sCountry, .sPhone, .sFAX, miEmpl, cType, miCompany)
        End With
        If miCompany = 0 Then
            miError = 0
            Exit Sub
        End If
    Case clCONTACTUPDATE + 1
    '===== Add the contact ====================
        With NewContact
            miContact = BLServer.AddContact(.sFirstName, .sLastName, .sTitle, miEmpl, .sPhone, .sSkill, .sTraining, miContact)
        End With
        If miContact = 0 Then
            miError = 0
            Exit Sub
        End If
End Select
End Sub
Private Sub BLServer_OnNewCustomer()
'MsgBox "OnNewCustomer Event", , "FAddItem-" & FMain.cboEmployees.Text
'    mbNewCompany = True
End Sub

Private Sub BLServer_OnUpdateDone()
'MsgBox "OnUpdateDone Event", , "FAddItem-" & FMain.cboEmployees.Text
'Debug.Print "BLServer event OnUpdateDone"
End Sub

Private Sub cboState_GotFocus()
    cboState.BackColor = &H80000005
    cboState.Refresh
End Sub

Private Sub cboState_LostFocus()
    If cboState.Text = "" Then cboState.BackColor = &HE0E0E0
    cboState.Refresh
End Sub

Private Sub chkTraining_Click(iIndex As Integer)
Dim iCounter As Integer
Static Inside As Integer

If Inside = 0 Then
Inside = 1

Select Case iIndex
    Case 0
        If chkTraining(iIndex).Value = vbChecked Then
            For iCounter = chkTraining.LBound + 1 To chkTraining.UBound
                chkTraining(iCounter).Value = vbUnchecked
            Next iCounter
        End If
    Case Else
        chkTraining(0) = vbUnchecked
End Select
Inside = 0
End If
End Sub

'##ModelId=3A0F61EB03DE
Private Sub cmdCancel_Click()
Me.ErrorCode = 9999
'    FMain.cmdEditContact.Enabled = True
'Unload Me
Me.Hide
End Sub
'##ModelId=3A0F61EC00A0
Private Sub cmdEnter_Click()
Dim iTemp As Long
Dim iCounter As Integer
Dim iCounter2 As Integer
Dim iItem As Integer
Static bInProcess As Boolean


On Error GoTo ErrorHandler
'    FMain.cmdEditContact.Enabled = True

'If bInProcess Then
'    Exit Sub
'End If

'Me.Hide

cmdEnter.Enabled = False

Select Case miAction
    Case clCOMPANYBYDATE + 1, clCOMPANYUPDATE + 1 'add company
    'Verify critical information is entered.
        If txtCompanyName.Text = "" Then
            MsgBox "Please enter a company name.", vbExclamation, "CL-FAddItem::Add Company"
            FAddItem.txtCompanyName.SetFocus
            cmdEnter.Enabled = True
            Exit Sub
        End If
        If txtPhone.Text = "" Then
            MsgBox "Please enter a phone number.", vbExclamation, "CL-FAddItem::Add Company"
            FAddItem.txtPhone.SetFocus
            cmdEnter.Enabled = True
            Exit Sub
        End If
        
        Set NewCompany = New CCustomer
        
        With NewCompany
            .sName = txtCompanyName.Text
            .sAddress = txtAddress.Text
            .sCity = txtCity.Text
            .sState = GetStateKey(cboState.ListIndex, True)
            .sZIP = txtZip.Text
            .sCountry = txtCountry.Text
            .sPhone = txtPhone.Text
            msPhone = .sPhone
            .sFAX = txtFax.Text
            .sType = mcType
            .ID = miCompany
        End With
        
        If miAction = (clCOMPANYBYDATE + 1) Then 'adding a contact
            miAction = clCONTACT + 1
            miContact = 10
            Form_Activate
            Exit Sub
        End If
        
    Case clCONTACT + 1, clCONTACTUPDATE + 1 'add contact
    'Verify critical information is entered.
        If txtFirstName.Text = "" Then
            MsgBox "Please enter the contact's name.", vbExclamation, "CL-FAddItem::Add Contact"
            FAddItem.txtFirstName.SetFocus
            cmdEnter.Enabled = True
            Exit Sub
        End If
        If txtContactNumber(0).Text = "" And txtContactNumber(1).Text = "" _
        And txtContactNumber(2).Text = "" And txtContactNumber(3).Text = "" Then
            MsgBox "Please enter a phone number.", vbExclamation, "CL-FAddItem::Add Contact"
            FAddItem.txtContactNumber(0).SetFocus
            cmdEnter.Enabled = True
            Exit Sub
        End If
        Set NewContact = New CContact
        
        'Select Training
        msTraining = ""
        For iCounter = chkTraining.LBound To chkTraining.UBound
            If chkTraining.Item(iCounter).Value = vbChecked Then
                msTraining = msTraining & Left(chkTraining.Item(iCounter).Caption, 1)
            End If
        Next iCounter
        
        'Get contact numbers
        'find the primary contact number
        For iCounter = optPrimary.LBound To optPrimary.UBound
            If optPrimary.Item(iCounter).Value = True Then Exit For
        Next iCounter
        
        'step through the phone number fields starting at the primary
        msPhone = ""
        For iCounter2 = iCounter To iCounter + 3
            If iCounter2 < 4 Then
                iItem = iCounter2
            Else
                iItem = iCounter2 - 4
            End If
            msPhone = msPhone & " (" & iItem & ") " & lblContactNumber.Item(iItem).Caption & txtContactNumber.Item(iItem).Text
        Next iCounter2
        
        With NewContact
            .sLastName = txtLastName.Text
'            If txtLastName.Text = "" Then .sLastName = "--" Else .sLastName = txtLastName.Text
            .sFirstName = txtFirstName.Text
'            If txtTitle.Text = "" Then .sTitle = "--" Else .sTitle = txtTitle.Text
            .sTitle = txtTitle.Text
            .sTraining = msTraining
            .sSkill = msSkill
            .sPhone = msPhone
            .ID = miContact
        End With
        
    Case clPRODUCT + 1 'add product
'        If Addproduct(txtProductName.Text, EmplID) = 0 Then
'            MsgBox "Error adding product", vbExclamation, "FAddItem"
'        End If
    Case clCALLTYPE + 1 'add call type
        If txtCallType.Text = "" Then
            MsgBox "Please enter the call type.", vbExclamation, "CL-FAddItem::Add CallType"
            FAddItem.txtCallType.SetFocus
            cmdEnter.Enabled = True
            Exit Sub
        End If
        
        Set NewCallCode = New CCallCode
        
        With NewCallCode
            .sName = txtCallType.Text
        End With
    Case Else
End Select

AddNewItems

Me.Hide
Exit Sub

ErrorHandler:
Me.ErrorCode = Err.Number
Me.Hide
End Sub
'##ModelId=3A0F61EC014A
Property Let EmplID(iData As Long)
    miEmpl = iData
End Property
'##ModelId=3A0F61EC0321
Property Get EmplID() As Long
    EmplID = miEmpl
End Property
'##ModelId=3A0F61ED003D
Property Let CompanyID(iData As Long)
    miCompany = iData
End Property
'##ModelId=3A0F61ED0250
Property Get CompanyID() As Long
    CompanyID = miCompany
End Property
'##ModelId=3A0F61ED0369
Property Let ContactID(iData As Long)
    miContact = iData
End Property
'##ModelId=3A0F61EE017F
Property Get ContactID() As Long
    ContactID = miContact
End Property
'##ModelId=3A0F61EE02AC
Property Let LinkID(iData As Long)
    miLink = iData
End Property
'##ModelId=3A0F61EF00D6
Property Get LinkID() As Long
    LinkID = miLink
End Property
'##ModelId=3A0F61EF0203
Property Let CodeID(iData As Long)
    miCode = iData
End Property
Property Get cType() As String
    cType = mcType
End Property
Property Let cType(sData As String)
    mcType = sData
End Property
'##ModelId=3A0F61F00042
Property Get CodeID() As Long
    CodeID = miCode
End Property
'##ModelId=3A0F61F0016E
Property Let ActionType(iData As Integer)
'1=company
'2=contact
'3=product
'4=call type

    miAction = iData
    iUpdate = iData
End Property
'##ModelId=3A0F61F0039F
Property Get ActionType() As Integer
    ActionType = miAction
End Property
'##ModelId=3A0F61F100ED
Property Let ErrorCode(iData As Integer)
    miError = iData
End Property
'##ModelId=3A0F61F1033C
Property Get ErrorCode() As Integer
    ErrorCode = miError
End Property
Property Let NewCompanyAdded(bData As Boolean)
    mbNewCompany = bData
End Property
Property Get NewCompanyAdded() As Boolean
    NewCompanyAdded = mbNewCompany
End Property
'##ModelId=3A0F61F20095
Property Let EntryText(sData As String)
    msText = sData
End Property
'##ModelId=3A0F61F202EE
Property Get EntryText() As String
    EntryText = msText
End Property

Private Sub Form_Activate()

Screen.MousePointer = vbHourglass
DisableFields

'If Me.EmplID = 0 Then
'    MsgBox "Must choose Employee Name before proceeding.", vbExclamation, "FAddItem"
'    cmdCancel_Click
'    Exit Sub
'End If
If Me.EmplID = 0 Then
    MsgBox "Must choose Employee Name before proceeding.", vbExclamation, "FAddItem"
    cmdCancel_Click
    Exit Sub
End If

miError = -1

If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    If BLServer Is Nothing Then
'        FMain.shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server, contact administrator.", vbCritical, "FAddItem::Load"
        Exit Sub
    End If
End If

Select Case miAction
    Case clCOMPANYBYDATE + 1, clCOMPANYUPDATE + 1, clNEWCOMPANIESNOCALLS + 1 'Add company
        cmdEnter.Enabled = True
        frmCompany.Left = 120
        frmCompany.Top = 0
        frmCompany.Visible = True
        frmContact.Visible = False
        frmProduct.Visible = False
        frmCallType.Visible = False
        
        If Not cboState.ListCount > 0 Then LoadStateCbo
            
        FillCompanyFields
    
    Case clCONTACT + 1, clCONTACTUPDATE + 1 'Add contact
        cmdEnter.Enabled = True
        frmContact.Left = 120
        frmContact.Top = 0
        frmCompany.Visible = False
        frmContact.Visible = True
        frmProduct.Visible = False
        frmCallType.Visible = False
        
        FillContactFields
    
    Case 3 'Add product
        cmdEnter.Enabled = True
        frmProduct.Left = 120
        frmProduct.Top = 0
        frmCompany.Visible = False
        frmContact.Visible = False
        frmProduct.Visible = True
        frmCallType.Visible = False
        txtProductName.Text = Me.EntryText
        
    Case 4 'Add call type
        cmdEnter.Enabled = True
        frmCallType.Left = 120
        frmCallType.Top = 0
        frmCompany.Visible = False
        frmContact.Visible = False
        frmProduct.Visible = False
        frmCallType.Visible = True
        txtCallType.Text = Me.EntryText
        
    Case Else
        LoadStateCbo
        frmCompany.Left = 120
        frmCompany.Top = 0
        frmCompany.Visible = True
        frmContact.Visible = False
        frmProduct.Visible = False
        frmCallType.Visible = False
        FillCompanyFields
End Select

EnableFields
Screen.MousePointer = Default

End Sub

Private Sub Form_Initialize()
    Me.cType = "U"

'Fill State Array
aState(50, 0) = "Wyoming"
aState(50, 1) = "WY"
aState(49, 0) = "Wisconsin"
aState(49, 1) = "WI"
aState(48, 0) = "West Virginia"
aState(48, 1) = "WV"
aState(47, 0) = "Washington"
aState(47, 1) = "WA"
aState(46, 0) = "Virginia"
aState(46, 1) = "VA"

aState(45, 0) = "Vermont"
aState(45, 1) = "VT"
aState(44, 0) = "Utah"
aState(44, 1) = "UT"
aState(43, 0) = "Texas"
aState(43, 1) = "TX"
aState(42, 0) = "Tennessee"
aState(42, 1) = "TN"
aState(41, 0) = "South Dakota"
aState(41, 1) = "SD"

aState(40, 0) = "South Carolina"
aState(40, 1) = "SC"
aState(39, 0) = "Rhode Island"
aState(39, 1) = "RI"
aState(38, 0) = "Ohio"
aState(38, 1) = "OH"
aState(37, 0) = "Pennsylvania"
aState(37, 1) = "PA"
aState(36, 0) = "Oregon"
aState(36, 1) = "OR"

aState(35, 0) = "Oklahoma"
aState(35, 1) = "OK"
aState(34, 0) = "North Dakota"
aState(34, 1) = "ND"
aState(33, 0) = "North Carolina"
aState(33, 1) = "NC"
aState(32, 0) = "New York"
aState(32, 1) = "NY"
aState(31, 0) = "New Mexico"
aState(31, 1) = "NM"

aState(30, 0) = "New Jersey"
aState(30, 1) = "NJ"
aState(29, 0) = "New Hampshire"
aState(29, 1) = "NH"
aState(28, 0) = "Nevada"
aState(28, 1) = "NV"
aState(27, 0) = "Nebraska"
aState(27, 1) = "NE"
aState(26, 0) = "Montana"
aState(26, 1) = "MT"

aState(25, 0) = "Missouri"
aState(25, 1) = "MO"
aState(24, 0) = "Mississippi"
aState(24, 1) = "MS"
aState(23, 0) = "Minnesota"
aState(23, 1) = "MN"
aState(22, 0) = "Michigan"
aState(22, 1) = "MI"
aState(21, 0) = "Massachusetts"
aState(21, 1) = "MA"

aState(20, 0) = "Maryland"
aState(20, 1) = "MD"
aState(19, 0) = "Maine"
aState(19, 1) = "ME"
aState(18, 0) = "Louisiana"
aState(18, 1) = "LA"
aState(17, 0) = "Kentucky"
aState(17, 1) = "KY"
aState(16, 0) = "Kansas"
aState(16, 1) = "KS"

aState(15, 0) = "Iowa"
aState(15, 1) = "IA"
aState(14, 0) = "Indiana"
aState(14, 1) = "IN"
aState(13, 0) = "Illinois"
aState(13, 1) = "IL"
aState(12, 0) = "Idaho"
aState(12, 1) = "ID"
aState(11, 0) = "Hawaii"
aState(11, 1) = "HI"

aState(10, 0) = "Georgia"
aState(10, 1) = "GA"
aState(9, 0) = "Florida"
aState(9, 1) = "FL"
aState(8, 0) = "Delaware"
aState(8, 1) = "DE"
aState(7, 0) = "Connecticut"
aState(7, 1) = "CT"
aState(6, 0) = "Colorado"
aState(6, 1) = "CO"

aState(5, 0) = "California"
aState(5, 1) = "CA"
aState(4, 0) = "Arkansas"
aState(4, 1) = "AR"
aState(3, 0) = "Arizona"
aState(3, 1) = "AZ"
aState(2, 0) = "Alaska"
aState(2, 1) = "AK"
aState(1, 0) = "Alabama"
aState(1, 1) = "AL"
aState(0, 0) = "Select State"
aState(0, 1) = "--"

LoadStateCbo

End Sub

'##ModelId=3A0F61F300AA
Private Sub Form_Load()

miError = -1
End Sub
'##ModelId=3A0F61F30190
Private Sub Form_Terminate()
    If Not BLServer Is Nothing Then Set BLServer = Nothing
    fMainForm.cmdEditContact.Enabled = True
    fMainForm.cmdNewContact.Enabled = True
    fMainForm.cmdNewCustomer.Enabled = True
End Sub

Private Sub optSkill_Click(Index As Integer)
    msSkill = Index + 1
End Sub

Private Sub optType_Click(Index As Integer)
Select Case Index
    Case 0 'end user
        Me.cType = "U"
    Case 1 'integrator
        Me.cType = "I"
    Case 2 'distributor
        Me.cType = "D"
    Case 3 'oem
        Me.cType = "O"
    Case 4 'service house
        Me.cType = "S"
    Case 5 'rep
        Me.cType = "R"
    Case 6 'Eurotherm drives company
        Me.cType = "E"
    Case 7 'Invensys sister company
        Me.cType = "F"
    Case Else
        Me.cType = "U"
End Select
End Sub
Private Sub FillCompanyFields()
Dim rs As ADODB.Recordset
Dim iIndex As Integer
Dim sTemp As String
Dim Cust As CCustomer

On Error GoTo FillCompanyFieldsErrorHandler

Set rs = BLServer.GetLbData2(clCOMPANYUPDATE, CStr(miCompany))
Set Cust = New CCustomer

With Cust
    If miAction = clCOMPANYUPDATE + 1 Then
        .ID = rs!ID
        .sPhone = rs!Phone
        .sFAX = rs!Fax
    End If
    .sName = rs!CompanyName
    .sAddress = rs!Address
    .sCity = rs!City
    .sState = rs!State
    .sZIP = rs!Zip
    .sCountry = rs!Country
    .sType = rs!Type

    txtCompanyName.Text = .sName
    txtAddress.Text = .sAddress
    txtCity.Text = .sCity
    iIndex = CInt(GetStateKey(.sState, False))
    cboState.ListIndex = iIndex
    sTemp = cboState.List(iIndex)
    txtZip.Text = .sZIP
    txtCountry.Text = .sCountry
    txtPhone.Text = .sPhone
    txtFax.Text = .sFAX
    optType(.iType).Value = True
End With
        
Set rs = Nothing
Exit Sub

FillCompanyFieldsErrorHandler:
Select Case Err.Number
    Case 94
    Case Else
End Select
Resume Next
Set rs = Nothing
End Sub
Private Sub FillContactFields()
Dim rs As ADODB.Recordset
Dim iPrimaryNumber As Integer
Dim iCounter As Integer
Dim iItem As Integer
Dim iFirstColon As Integer
Dim iEndOfNumber As Integer
Dim iNextItem As Integer
Dim iLengthOfNumber As Integer

Dim Cont As CContact

On Error GoTo FillContactFieldsErrorHandler

Set rs = BLServer.GetLbData2(clCONTACTUPDATE, CStr(miContact))
Set Cont = New CContact

With Cont
    If Not (miAction - 1 = clCONTACT) Then 'must be editing current contact
        .sFirstName = rs!FirstName
        .sLastName = rs!LastName
        .sTitle = rs!Title
        .sSkill = rs!Skill
        .sTraining = rs!Training
        .sPhone = rs!Phone
    ElseIf rs!ID <> 10 Then 'new contact at existing location
        .sPhone = rs!Phone
    End If
'    .sName = "" 'rs!sName
'    .sPhone = rs!Phone
'    miContact = rs!ID

    txtFirstName.Text = .sFirstName
    txtLastName.Text = .sLastName
    txtTitle.Text = .sTitle
    msTraining = .sTraining
    
    For iCounter = chkTraining.LBound To chkTraining.UBound
        If InStr(1, msTraining, Left(chkTraining.Item(iCounter).Caption, 1)) > 0 Then
            chkTraining.Item(iCounter).Value = vbChecked
        End If
    Next iCounter
    
    msSkill = .sSkill
    optSkill(msSkill - 1).Value = True
    
    txtContactNumber(0).Text = msPhone
    txtContactNumber(1).Text = ""
    txtContactNumber(2).Text = ""
    txtContactNumber(3).Text = ""
    
    msPhone = .sPhone

    optPrimary(CInt(Mid(.sPhone, 2, 1))).Value = True
    'step through the phone number fields starting at the primary
    For iCounter = iPrimaryNumber To iPrimaryNumber + 3
        If iCounter < 4 Then
            iItem = iCounter
        Else
            iItem = iCounter - 4
        End If
        If InStr(1, msPhone, "(" & iItem & ")") = 0 Then Exit For
            iFirstColon = InStr(InStr(1, msPhone, "(" & iItem & ")"), msPhone, ":") + 1
        iNextItem = iItem + 1
        If iNextItem > 3 Then iNextItem = iNextItem - 4
        iEndOfNumber = InStr(iFirstColon, msPhone, "(" & iNextItem & ")")
        If iEndOfNumber = 0 Then iEndOfNumber = Len(msPhone) + 1
        iLengthOfNumber = iEndOfNumber - iFirstColon
        txtContactNumber(iItem).Text = Trim(Mid(msPhone, iFirstColon, iLengthOfNumber))
    Next iCounter
End With

'miContact = rs!ID
Set rs = Nothing
Exit Sub

FillContactFieldsErrorHandler:
Set rs = Nothing

End Sub
Sub LoadStateCbo()
Dim iCounter As Integer

For iCounter = 50 To 0 Step -1
    cboState.AddItem aState(iCounter, 0), 0
    cboState.ItemData(cboState.NewIndex) = iCounter
Next
End Sub
Function GetStateKey(sIndex As String, fGetGive As Boolean) As String

If sIndex = "" And Not fGetGive Then sIndex = "--"

If fGetGive Then
    GetStateKey = aState(sIndex, 1)
Else
    Dim iCounter As Integer
    For iCounter = 0 To 50
        If aState(iCounter, 1) = sIndex Then
            GetStateKey = iCounter
        End If
    Next
End If

'cboState.Sorted

End Function

Private Sub txtAddress_GotFocus()
    txtAddress.BackColor = &H80000005
    txtAddress.SelStart = 0
    txtAddress.SelLength = Len(txtAddress.Text)
End Sub

Private Sub txtAddress_LostFocus()
    If txtAddress.Text = "" Then txtAddress.BackColor = &HE0E0E0
End Sub

Private Sub txtCity_GotFocus()
    txtCity.BackColor = &H80000005
    txtCity.SelStart = 0
    txtCity.SelLength = Len(txtCity.Text)
End Sub

Private Sub txtCity_LostFocus()
    If txtCity.Text = "" Then txtCity.BackColor = &HE0E0E0
End Sub

Private Sub txtCompanyName_GotFocus()
    txtCompanyName.BackColor = &H80000005
    txtCompanyName.SelStart = 0
    txtCompanyName.SelLength = Len(txtCompanyName.Text)
End Sub

Private Sub txtContactNumber_GotFocus(Index As Integer)
    txtContactNumber(Index).BackColor = &H80000005
    txtContactNumber(Index).SelStart = 0
    txtContactNumber(Index).SelLength = Len(txtContactNumber(Index).Text)
End Sub

Private Sub txtCountry_GotFocus()
    txtCountry.BackColor = &H80000005
    txtCountry.SelStart = 0
    txtCountry.SelLength = Len(txtCountry.Text)
End Sub

Private Sub txtCountry_LostFocus()
    If txtCountry.Text = "" Then txtCountry.BackColor = &HE0E0E0
End Sub
Private Sub txtFax_GotFocus()
    txtFax.BackColor = &H80000005
    txtFax.SelStart = 0
    txtFax.SelLength = Len(txtFax.Text)
End Sub

Private Sub txtFax_LostFocus()
    If txtFax.Text = "" Then txtFax.BackColor = &HE0E0E0

End Sub

Private Sub txtFirstName_GotFocus()
    txtFirstName.BackColor = &H80000005
    txtFirstName.SelStart = 0
    txtFirstName.SelLength = Len(txtFirstName.Text)
End Sub

Private Sub txtLastName_GotFocus()
    txtLastName.BackColor = &H80000005
    txtLastName.SelStart = 0
    txtLastName.SelLength = Len(txtLastName.Text)
End Sub

Private Sub txtLastName_LostFocus()
    If txtLastName.Text = "" Then txtLastName.BackColor = &HE0E0E0
End Sub

Private Sub txtPhone_GotFocus()
    txtPhone.BackColor = &H80000005
    txtPhone.SelStart = 0
    txtPhone.SelLength = Len(txtPhone.Text)
End Sub

Private Sub txtTitle_GotFocus()
    txtTitle.BackColor = &H80000005
    txtTitle.SelStart = 0
    txtTitle.SelLength = Len(txtTitle.Text)
End Sub

Private Sub txtTitle_LostFocus()
    If txtTitle.Text = "" Then txtTitle.BackColor = &HE0E0E0
End Sub

Private Sub txtZip_GotFocus()
    txtZip.BackColor = &H80000005
    txtZip.SelStart = 0
    txtZip.SelLength = Len(txtZip.Text)
End Sub

Private Sub txtZip_LostFocus()
    If txtZip.Text = "" Then txtZip.BackColor = &HE0E0E0
End Sub

Private Sub EnableFields()
    frmCompany.Enabled = True
    frmContact.Enabled = True
    frmCallType.Enabled = True
    frmProduct.Enabled = True
    
    txtAddress.BackColor = &HE0E0E0
    txtCity.BackColor = &HE0E0E0
    cboState.BackColor = &HE0E0E0
    txtZip.BackColor = &HE0E0E0
    txtCountry.BackColor = &HE0E0E0
    txtFax.BackColor = &HE0E0E0
    
    txtLastName.BackColor = &HE0E0E0
    txtTitle.BackColor = &HE0E0E0
    txtContactNumber(1).BackColor = &HE0E0E0
    txtContactNumber(2).BackColor = &HE0E0E0
    txtContactNumber(3).BackColor = &HE0E0E0

End Sub
Private Sub DisableFields()
    frmCompany.Enabled = False
    frmContact.Enabled = False
    frmCallType.Enabled = False
    frmProduct.Enabled = False
End Sub

