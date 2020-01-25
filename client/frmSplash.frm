VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7455
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   4590
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   7380
      Begin VB.Timer tmrProgressTimer 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   6840
         Top             =   4080
      End
      Begin ComctlLib.ProgressBar pgbLoadTVW 
         Height          =   200
         Left            =   240
         TabIndex        =   11
         Top             =   540
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   344
         _Version        =   327682
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1
      End
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   2880
         Picture         =   "frmSplash.frx":0442
         ScaleHeight     =   1005
         ScaleWidth      =   1515
         TabIndex        =   2
         Top             =   1200
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2760
         TabIndex        =   9
         Tag             =   "Product"
         Top             =   2160
         Width           =   1785
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LicenseTo"
         Height          =   255
         Left            =   270
         TabIndex        =   1
         Tag             =   "LicenseTo"
         Top             =   300
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2160
         TabIndex        =   8
         Tag             =   "CompanyProduct"
         Top             =   720
         Width           =   3000
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5865
         TabIndex        =   7
         Tag             =   "Platform"
         Top             =   2400
         Width           =   1140
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6075
         TabIndex        =   6
         Tag             =   "Version"
         Top             =   2760
         Width           =   930
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Warning"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Tag             =   "Warning"
         Top             =   4320
         Width           =   6975
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         Tag             =   "Company"
         Top             =   3360
         Width           =   3375
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         Height          =   255
         Left            =   3750
         TabIndex        =   4
         Tag             =   "Copyright"
         Top             =   3120
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC8A0001F"


'##ModelId=39EBC8A000D3
Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    lblCompanyProduct.Caption = "Big Dog Tools"
    lblCompany.Caption = "Big Dog, LLC."
    lblCopyright.Caption = "All rights reserved.  Copyright 2000."
    lblLicenseTo.Caption = "BSA"
    lblPlatform.Caption = "NT 4"
    lblWarning.Caption = "This program is for the use of my friends. Don't take it or else!"
End Sub

Private Sub tmrProgressTimer_Timer()
If LoadProgressActive Then
    Me.pgbLoadTVW.Value = LoadProgress
    Me.pgbLoadTVW.Max = LoadProgressMax
    If LoadProgress = Me.pgbLoadTVW.Max Then LoadProgressActive = False
End If
End Sub
