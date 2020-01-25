VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FMain 
   Caption         =   "Call Tracker"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12120
   FillColor       =   &H80000013&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "FProdSupLoading.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8715
   ScaleWidth      =   12120
   Begin VB.Timer tmrBLStatus 
      Interval        =   100
      Left            =   5520
      Top             =   120
   End
   Begin VB.CommandButton cmdForceTVWLoad 
      Caption         =   "Reload"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10380
      TabIndex        =   51
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox txtMinutes 
      Alignment       =   2  'Center
      DataSource      =   "datProdSupLoading"
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
      TabIndex        =   0
      Text            =   "Enter Minutes Here"
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10380
      TabIndex        =   49
      Top             =   5700
      Width           =   1575
   End
   Begin VB.CommandButton cmdTimer 
      BackColor       =   &H00FF0000&
      Caption         =   "Start"
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
      Left            =   6600
      MaskColor       =   &H00FFFF00&
      TabIndex        =   48
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdNewCustomer 
      Caption         =   "New Company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10380
      TabIndex        =   47
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdNewContact 
      Caption         =   "New Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10380
      TabIndex        =   46
      ToolTipText     =   "Add a new contact to the selected company."
      Top             =   3900
      Width           =   1575
   End
   Begin VB.CommandButton cmdEditContact 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10380
      TabIndex        =   45
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtEnterCaseID 
      Alignment       =   2  'Center
      DataSource      =   "datProdSupLoading"
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
      Left            =   10380
      TabIndex        =   43
      Text            =   "Enter ID"
      Top             =   585
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Enter  Call"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10380
      TabIndex        =   19
      Top             =   1140
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6000
      Top             =   120
   End
   Begin VB.CommandButton cmdShowMore 
      Caption         =   "More >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   10380
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   39
      Top             =   8460
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "8/24/01"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "12:21 AM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   13053
            MinWidth        =   13053
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   0
            Text            =   "FMain"
            TextSave        =   "FMain"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmCall 
      Height          =   2235
      Left            =   120
      TabIndex        =   28
      Top             =   660
      Visible         =   0   'False
      Width           =   10215
      Begin VB.TextBox txtCallNote 
         Height          =   1635
         Left            =   4320
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   480
         Width           =   5775
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   1935
      End
      Begin VB.ListBox lstItem 
         Height          =   1230
         Index           =   2
         Left            =   120
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   30
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstItem 
         Height          =   1230
         Index           =   3
         Left            =   2160
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lblCallNote 
         Caption         =   "Note"
         Height          =   255
         Left            =   4320
         TabIndex        =   37
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblProduct 
         Caption         =   "Product"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblCallCode 
         Caption         =   "Call Code"
         Height          =   255
         Left            =   2160
         TabIndex        =   33
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   10380
      TabIndex        =   27
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame frmCustomer 
      Height          =   5535
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Visible         =   0   'False
      Width           =   10215
      Begin ComctlLib.TreeView tvwCustomers 
         DragIcon        =   "FProdSupLoading.frx":0442
         Height          =   4935
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   8705
         _Version        =   327682
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imgTreeViewIcons"
         Appearance      =   1
      End
      Begin VB.TextBox txtCallHistory 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   4320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   480
         Width           =   5775
      End
      Begin VB.ListBox lstItem 
         Height          =   1425
         Index           =   0
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   4095
      End
      Begin VB.ListBox lstItem 
         Height          =   1230
         Index           =   1
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3000
         Width           =   4095
      End
      Begin VB.Label lblTVKey 
         Height          =   255
         Left            =   5640
         TabIndex        =   50
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Call History"
         Height          =   255
         Left            =   4320
         TabIndex        =   40
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblContact 
         Caption         =   "Contact Name"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblCustomer 
         Caption         =   "Customer Name"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdDelCall 
      Caption         =   "Del"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton cmdDelCall 
      Caption         =   "Del"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton cmdDelCall 
      Caption         =   "Del"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdDelCall 
      Caption         =   "Del"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdDelCall 
      Caption         =   "Del"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   495
   End
   Begin VB.ComboBox cboEmployees 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   100
      Width           =   4215
   End
   Begin VB.VScrollBar vsbChangeDate 
      Height          =   495
      Left            =   4080
      Max             =   0
      Min             =   30
      TabIndex        =   3
      Top             =   705
      Width           =   255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Done - Finally!"
      Enabled         =   0   'False
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
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   4335
   End
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
      Index           =   0
      Left            =   480
      TabIndex        =   12
      Top             =   4920
      Width           =   2295
   End
   Begin VB.VScrollBar vscCalls 
      Height          =   2295
      Left            =   4800
      Max             =   500
      TabIndex        =   11
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton cmdShowMore 
      Caption         =   "More >>"
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
      Index           =   0
      Left            =   2760
      TabIndex        =   36
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Shape shBLStatus 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   11760
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
   Begin ComctlLib.ImageList imgTreeViewIcons 
      Left            =   8400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   42
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":0884
            Key             =   "D"
            Object.Tag             =   "D"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":0B9E
            Key             =   "E"
            Object.Tag             =   "E"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":0EB8
            Key             =   "F"
            Object.Tag             =   "F"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":11D2
            Key             =   "I"
            Object.Tag             =   "I"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":14EC
            Key             =   "O"
            Object.Tag             =   "O"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":1806
            Key             =   "R"
            Object.Tag             =   "R"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":1B20
            Key             =   "S"
            Object.Tag             =   "S"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":1E3A
            Key             =   "U"
            Object.Tag             =   "U"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":2154
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":246E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":2788
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":2AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":2DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":30D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":33F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":370A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":3A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":3D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":4058
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":4372
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":468C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":49A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":4CC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":4FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":52F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":560E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":5928
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":5C42
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":5F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":6276
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":6590
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":68AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":6BC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":6EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":71F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":7512
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":782C
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":7B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":7E60
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":817A
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":8494
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading.frx":87AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Case ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   42
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblCaseID 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10380
      TabIndex        =   41
      ToolTipText     =   "Double-click to enter Case ID"
      Top             =   585
      Width           =   1575
   End
   Begin VB.Label lblEntryData 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   4
      Left            =   720
      TabIndex        =   10
      Top             =   4440
      Width           =   4095
   End
   Begin VB.Label lblEntryData 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   3
      Left            =   720
      TabIndex        =   9
      Top             =   3960
      Width           =   4095
   End
   Begin VB.Label lblEntryData 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   8
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Label lblEntryData 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Label lblEntryData 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label lblTotalCalls 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC85C0112"
Option Explicit

'##ModelId=39EBC85C01AA
Dim Employees As CEmployees
'##ModelId=39EBC85C01BE
Dim Customers As CCustomers
'##ModelId=39EBC85C01D0
Dim Contacts As CContacts
'##ModelId=39EBC85C01DC
Dim Products As CProducts
'##ModelId=39EBC85C01EE
Dim CallCodes As CCallCodes
'##ModelId=39EBC85C01F8
Dim Entries As Collection
'##ModelId=39EBC85C0240
Dim Entry As CEntry
'##ModelId=39EBC85C02AC
Dim colLinks As CLinks
'##ModelId=39EBC85C02B8
Dim Calls As CCalls

'##ModelId=39EBC85C02CA
Dim Number As Integer
'##ModelId=39EBC85C031A
Dim EmplID As Long
'##ModelId=39EBC85C0374
Dim ErrorCondition As Integer
Dim rsGeneric As ADODB.Recordset

Const clCOMPANY As Integer = 0
Const clCONTACT As Integer = 1
Const clPRODUCT As Integer = 2 ''
Const clCALLCODE As Integer = 3 ''
Const clCALL As Integer = 4
Const clLINK As Integer = 5
Const clHISTORY As Integer = 6
Const clCALLTYPE As Integer = 7
Const clEMPLOYEE As Integer = 8
Const clCOMPANYUPDATE As Integer = 9
Const clCONTACTUPDATE As Integer = 10
Const clCONTACT2 As Integer = 12
Const clCOMPANY2 As Integer = 13
Const clCONTACT3 As Integer = 14
Const clCASEID As Integer = 15
Const clCOMPANY4 As Integer = 16
Const clNEWCALLS As Integer = 17
Const clNEWCOMPANIES As Integer = 18
Const clNEWCONTACTS As Integer = 19

Const clNODROP As Long = 29
Const clSELECTED As Long = 30

Const ceLONG As Integer = 1
Const ceSTRING As Integer = 2

Dim StartTime As Date
Dim iSelectedContactID As Long
Dim iSelectedCompanyID As Long
Dim iNewCompanyID As Long
Dim fDrag As Boolean

Dim iLastAction As Long
Dim iLastCompany As Long
Dim iLastContact As Long
Dim sCurrentParent As String
Dim dLastCompanyUpdate As Date

'##ModelId=3A0F61DC0025
Dim WithEvents BLServer As CallTrackerBLServer.BLServer
Attribute BLServer.VB_VarHelpID = -1

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const TVM_GETCOUNT = &H1105&
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                        (ByVal hwnd As Long, _
                        ByVal wMsg As Long, ByVal _
                        wParam As Long, _
                        lParam As Any) As Long
'##ModelId=3A0F61DC002F
Function InitLists() As Integer
On Error GoTo InitListsErrorHandler

Set Employees = New CEmployees
Set Products = New CProducts
Set CallCodes = New CCallCodes

'Connect to the BL Server
frmSplash.Label1.Caption = "Initializing BL Server ..."
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    shBLStatus.FillColor = vbGreen
    If BLServer Is Nothing Then
        shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server, contact administrator." & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::InitLists"
        Exit Function
    End If
End If

'Fill the local collections
frmSplash.Label1.Caption = "Getting Products ..."
LoadLists clPRODUCT, Products
frmSplash.Label1.Caption = "Getting Call Codes ..."
LoadLists clCALLCODE, CallCodes
frmSplash.Label1.Caption = "Getting Employees ..."
LoadLists clEMPLOYEE, Employees

'Load the listboxes and comboboxes
frmSplash.Label1.Caption = "Loading Lists ..."
If LoadComboBox(cboEmployees, Employees) = 0 Then MsgBox "error"
If LoadListBox(lstItem(clPRODUCT), Products) = 0 Then MsgBox "Couldn't load listbox", , "BLogic-GetCustomers"
If LoadListBox(lstItem(clCALLCODE), CallCodes) = 0 Then MsgBox "Couldn't load listbox", , "BLogic-GetCustomers"

'Load the customer treeview
frmSplash.Label1.Caption = "Loading Companies ..."
dLastCompanyUpdate = Now()
'LoadAllCustomers clCOMPANY

'Set Employees = Nothing
Set Products = Nothing
Set CallCodes = Nothing
Set BLServer = Nothing
shBLStatus.FillColor = vbRed

InitLists = 0 'Err.Number

Exit Function

InitListsErrorHandler:
MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::InitLists"
End Function
'##ModelId=3A0F61DC00B1
Sub ReInit(Index As Integer)
'============================================================
'This sub reinitializes the collections and list boxes after
'adding a new record for the list (like new customer and contact)
'============================================================

If Index = -1 Then Exit Sub

'Clear the listbox getting new information
Select Case Index
    Case clCOMPANY
'        BLServer_OnNewCustomer
        LoadOneCustomer clCOMPANY, Me.LastCompany 'load parents
        LoadOneCustomer clCONTACT, Me.LastContact 'load children
    Case clCONTACT
        LoadOneCustomer clCONTACT, Me.LastContact 'load children
    Case clPRODUCT
'        lstItem(Index).Clear
    Case clCALLCODE
        lstItem(Index).Clear
        Set CallCodes = New CCallCodes
        LoadLists clCALLCODE, CallCodes
        If LoadListBox(lstItem(clCALLCODE), CallCodes) = 0 Then MsgBox "Couldn't load listbox", vbCritical, "FMain::ReInit"
        GoToItem lstItem(clCALLCODE), FAddItem.CodeID
        Set CallCodes = Nothing
    Case clCOMPANYUPDATE
        LoadOneCustomer clCOMPANY2, Me.LastCompany 'load parents
    Case clCONTACTUPDATE
        LoadOneCustomer clCONTACT2, Me.LastContact 'load children
End Select

'Unload FAddItem
End Sub
'##ModelId=3A0F61DC01D4
Sub InitForm()
frmSplash.Label1.Caption = "Initializing ..."

Number = 0

'Initialize Employee combo box
cboEmployees.Clear
Set Entries = New Collection

Me.Width = 5295
Me.Height = 6150
sCurrentParent = ""
Me.LastAction = clDEFAULTLASTACTION
Me.LastCompany = clDEFAULTLASTCOMPANY
Me.LastContact = clDEFAULTLASTCONTACT

cmdShowMore_Click (1) 'resizes the form, will delete this in the future
Load FAddItem 'gets the form into memory for faster adds

'set date
lblDate.Caption = Format(Now(), "medium date")
End Sub
Sub AddItem(iIndex As Integer)
Dim iCompanyKey As Long

'-------------------------------------------------------
'Initialize
Me.LastCompany = clDEFAULTLASTCOMPANY 'Default for no company
Me.LastAction = iIndex

'Get the company id
If Not tvwCustomers.SelectedItem Is Nothing Then
    With tvwCustomers.SelectedItem
    'Make sure a contact node is selected and get the parent
        If Left(.Key, 1) = "c" Then 'its a child node
            Me.LastCompany = CLng(Right(.Parent.Key, Len(.Parent.Key) - 1))
        ElseIf Left(.Key, 1) = "p" Then 'its a parent node
            Me.LastCompany = CLng(Right(.Key, Len(.Key) - 1))
        End If
    End With
End If

'-------------------------------------------------------
'If updating contact, get the contact id
'Get the contact id
If Not tvwCustomers.SelectedItem Is Nothing Then
    With tvwCustomers.SelectedItem
    'Make sure a contact node is selected and get the parent
        If Left(.Key, 1) = "c" Then 'its a child node
            Me.LastContact = CLng(Right(.Key, Len(.Key) - 1))
        
        ElseIf .Children > 0 Then
            Me.LastContact = CLng(Right(.Child.FirstSibling.Key, Len(.Child.FirstSibling.Key) - 1))
        Else
            Me.LastContact = clDEFAULTLASTCONTACT
        End If
    End With
End If

With FAddItem

    .CompanyID = Me.LastCompany
    .ContactID = Me.LastContact
    .EmplID = cboEmployees.ItemData(cboEmployees.ListIndex)
    .ActionType = iIndex + 1
    
    .Show vbModal
        
    Me.LastCompany = .CompanyID
    Me.LastContact = .ContactID
    
    FAddItem.Hide
    
    cmdEditContact.Enabled = True
    cmdNewContact.Enabled = True
    cmdNewCustomer.Enabled = True

    If .ErrorCode = 9999 Then 'cancelled edit/add
        Exit Sub
    ElseIf .ErrorCode <> 0 Then 'edit/add succeeded
        ReInit iIndex
    Else 'failed the add/edit
        MsgBox "Error adding item, error = " & .ErrorCode, vbCritical, "CL-Main::AddItem"
    End If
End With
End Sub

Private Sub BLServer_OnNewCustomer()
'MsgBox "OnNewCustomer Event", , FMain.cboEmployees.Text
If FAddItem.NewCompanyAdded = False Then
    LoadOneCustomer clCOMPANY4, CLng(dLastCompanyUpdate) 'load parents
'    LoadOneCustomer clCONTACT, Me.LastContact 'load children
End If
'    LoadOneCustomer clCOMPANY, Me.LastCompany 'load parents
'    LoadOneCustomer clCONTACT, Me.LastContact 'load children
dLastCompanyUpdate = Now()
End Sub

'##ModelId=3A0F61DC0242
Private Sub cboEmployees_Click()
Dim vCounter As Variant

For Each vCounter In Employees
    If vCounter.sName = cboEmployees.Text Then
        EmplID = vCounter.ID
        Exit For
    End If
Next vCounter
If cboEmployees.ListIndex = 0 Then
    cmdNewContact.Enabled = False
    cmdNewCustomer.Enabled = False
    cmdEditContact.Enabled = False
    cmdUpdate.Enabled = False
Else
'    cmdNewContact.Enabled = True
    If Not tvwCustomers.SelectedItem Is Nothing Then
        tvwCustomers.SelectedItem = Nothing
    End If
    cmdNewCustomer.Enabled = True
'    cmdEditContact.Enabled = True
    cmdUpdate.Enabled = True
End If
End Sub
'##ModelId=3A0F61DC02C4
Private Sub cmdCancel_Click(Index As Integer)
Unload Me
End Sub
'##ModelId=3A0F61DD001C
Private Sub cmdDelCall_Click(Index As Integer)
Dim EItem As Integer
Dim EntryText As String
Dim MaxItems As Integer

Number = Number - 1
If Number < 0 Then Number = 0
lblTotalCalls.Caption = "Total Calls Entered = " & CStr(Number)

If Entries.Count > 0 Then

    EItem = CInt(Left(lblEntryData(Index).Caption, InStr(lblEntryData(Index).Caption, " ") - 1))

    Entries.Remove CStr(EItem)

    If Entries.Count > 5 Then MaxItems = 4 Else MaxItems = Entries.Count - 1

    'disable all del buttons and clear list
    For EItem = 0 To 4
        cmdDelCall(EItem).Enabled = False
        lblEntryData(EItem).Caption = ""
    Next
        
    'enable del buttons
    For EItem = 0 To MaxItems
        cmdDelCall(EItem).Enabled = True
    Next
    
    'fill list
    For EItem = 0 To MaxItems
        Set Entry = Entries.Item(Entries.Count - EItem)
        EntryText = Entry.EIndex & " : " & Entry.EmplID & " : " & Entry.EDate & " : " & Entry.ETime
        If Entries.Count >= EItem + 1 Then lblEntryData(EItem).Caption = EntryText
        If Entries.Count - 1 <= EItem Then Exit For
    Next
        
    If Entries.Count > 5 Then MaxItems = 4 Else MaxItems = Entries.Count - 1
        
    For EItem = 0 To MaxItems
        cmdDelCall(EItem).Enabled = True
        If Entries.Count - 1 <= EItem Then Exit For
    Next
    
End If

End Sub
'##ModelId=3A0F61DD015D
Private Sub cmdExit_Click()
Dim msg As String
'Dim Entry As New CEntry
Dim BLError As Long
Dim oEntries As Object
'
Set oEntries = Entries

On Error Resume Next    ' Enable error trapping.

BLError = BLServer.AddEntries(oEntries)
MsgBox BLError & vbCrLf & BLServer.Status
If BLError <> 0 Then
    MsgBox "Failed to add calls" & vbCrLf & "Error number: " & BLError & " from BLServer"
End If
cmdExit.Caption = "Finished, Closing ..."

If Not BLServer Is Nothing Then Set BLServer = Nothing

Unload Me

End Sub

Private Sub cmdEditContact_Click()
'edit company or contact
'check the selected item

cmdNewCustomer.Enabled = False
cmdNewContact.Enabled = False
cmdEditContact.Enabled = False
Screen.MousePointer = vbHourglass

With tvwCustomers
    If Not .SelectedItem Is Nothing Then
        If Left(.SelectedItem.Key, 1) = "p" Then
            AddItem clCOMPANYUPDATE
        ElseIf Left(.SelectedItem.Key, 1) = "c" Then
            AddItem clCONTACTUPDATE
        End If
    End If
End With
End Sub

Private Sub cmdForceTVWLoad_Click()

cmdEditContact.Enabled = False
cmdNewContact.Enabled = False
txtCallHistory.Text = ""
sCurrentParent = ""
tvwCustomers.Nodes.Clear
tvwCustomers.Refresh

LoadAllCustomers clCOMPANY

End Sub

Private Sub cmdNewContact_Click()
Screen.MousePointer = vbHourglass
cmdNewCustomer.Enabled = False
cmdNewContact.Enabled = False
cmdEditContact.Enabled = False
AddItem (clCONTACT)
End Sub

Private Sub cmdNewCustomer_Click()
On Error GoTo NoCustomer

Screen.MousePointer = vbHourglass
cmdNewCustomer.Enabled = False
cmdNewContact.Enabled = False
cmdEditContact.Enabled = False
AddItem (clCOMPANY)
tvwCustomers.Nodes.Item("p" & Me.LastCompany).Selected = True
tvwCustomers.Nodes.Item("p" & Me.LastCompany).Expanded = True

NoCustomer:
Select Case Err.Number
    Case 35601
        tvwCustomers.Nodes.Item(1).Selected = True
        Me.LastCompany = Right(tvwCustomers.Nodes.Item(1).Key, Len(tvwCustomers.Nodes.Item(1).Key) - 1)
    Case Else
End Select
Resume Next
End Sub

Private Sub cmdQuery_Click()
Load fQuery
Set fQuery.Products = lstItem(2)
Set fQuery.CallCodes = lstItem(3)
fQuery.Show vbModal
End Sub

'##ModelId=3A0F61DD01DF
Private Sub cmdShowMore_Click(Index As Integer)
If cmdShowMore(Index).Caption = "More >>" Then
    frmCall.Visible = True
    frmCustomer.Visible = True
    cmdUpdate.Visible = True
    cmdCancel(1).Visible = True
    cmdShowMore(0).Caption = "<< Less"
    cmdShowMore(1).Caption = "<< Less"
    cmdShowMore(1).Visible = True
    Me.Width = 12120
    Me.Height = 9120
    txtMinutes.Left = 4440
    txtMinutes.Top = 100
    lblDate.Left = 8400
    lblDate.Top = 120
    vsbChangeDate.Left = 10080
    vsbChangeDate.Top = 100
Else
    frmCustomer.Visible = False
    frmCall.Visible = False
    cmdUpdate.Visible = False
    cmdCancel(1).Visible = False
    cmdShowMore(0).Caption = "More >>"
    cmdShowMore(1).Caption = "More >>"
    cmdShowMore(1).Visible = False
    Me.Width = 5295
    Me.Height = 6150
    txtMinutes.Left = 120
    txtMinutes.Top = 720
    lblDate.Left = 8400
    lblDate.Top = 120
    vsbChangeDate.Left = 10080
    vsbChangeDate.Top = 100
End If
End Sub

Private Sub cmdTimer_Click()
Static iTimerOn As Integer
If iTimerOn = 1 Then 'timer ON
    tmrMain.Enabled = False
    cmdTimer.Caption = "Start"
    iTimerOn = 0
    StartTime = 0
Else 'timer OFF
    tmrMain.Enabled = True
    cmdTimer.Caption = "Stop"
    If txtMinutes.Text = "Enter Minutes Here" Or txtMinutes.Text = "" Then txtMinutes.Text = "0"
    StartTime = Now() - CDate(txtMinutes.Text)
    iTimerOn = 1
End If
End Sub

'##ModelId=3A0F61DD0329
Private Sub cmdUpdate_Click()
Dim sText As String
Dim iCustomerID As Long
Dim iContactID As Long
Dim iProductID As Long
Dim iCallCodeID As Long
Dim iEmplID As Long
Dim vCount As Variant
Dim iCallTime As Integer

On Error GoTo ErrorHandler

If tvwCustomers.SelectedItem.Parent Is Nothing Then
    MsgBox "Please select contact.", vbExclamation, "FMain::cmdUpdate"
    Exit Sub
End If
iCustomerID = Right(tvwCustomers.SelectedItem.Parent.Key, Len(tvwCustomers.SelectedItem.Parent.Key) - 1)
iContactID = Right(tvwCustomers.SelectedItem.Key, Len(tvwCustomers.SelectedItem.Key) - 1)
If iContactID = 0 Then
    MsgBox "Please select contact.", vbExclamation, "FMain::cmdUpdate"
    Exit Sub
End If

If lstItem(2).ListIndex = -1 Then
    MsgBox "Please select product.", vbExclamation, "FMain::cmdUpdate"
    Exit Sub
End If
iProductID = lstItem(2).ItemData(lstItem(2).ListIndex)

If lstItem(3).ListIndex = -1 Then
    MsgBox "Please select call code.", vbExclamation, "FMain::cmdUpdate"
    Exit Sub
End If
iCallCodeID = lstItem(3).ItemData(lstItem(3).ListIndex)

iEmplID = cboEmployees.ItemData(cboEmployees.ListIndex)

iCallTime = 6

If tmrMain.Enabled = True Then tmrMain.Enabled = False

If Not txtMinutes.Text = "Enter Minutes Here" And Not txtMinutes.Text = "" Then
    If CLng(CDate(txtMinutes.Text) * 24 * 60) > 480 Then
        iCallTime = CInt(txtMinutes.Text)
    Else
        iCallTime = CInt(CDate(txtMinutes.Text) * 24 * 60)
    End If
Else
    MsgBox "Please enter the call time.", vbExclamation, "FMain::cmdUpdate"
    txtMinutes.Text = "Enter Minutes Here"
    Exit Sub
End If

If txtCallNote.Text = "" Then
    MsgBox "Enter call notes.", vbExclamation, "FMain::cmdUpdate"
    Exit Sub
End If

'Connect to the BL Server

If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    shBLStatus.FillColor = vbGreen
    If BLServer Is Nothing Then
        shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server, contact administrator." & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "FMain::cmdUpdate"
        Exit Sub
    End If
End If

If BLServer.AddCall(iCustomerID, iContactID, iCallCodeID, iProductID, iEmplID, lblDate.Caption, txtCallNote.Text, iCallTime, CLng(lblCaseID.Caption)) = 0 Then
    MsgBox "Error adding call", vbCritical, "FMain::cmdUpdate"
    Exit Sub
End If

txtCallHistory.Text = sText
GetCompanyHistory iCustomerID, txtCallHistory
ResetControls

Set BLServer = Nothing

Exit Sub

ErrorHandler:
    MsgBox "Failed to add call to database : " & Err.Description, vbCritical, "CL-cmdUpdate::CLICK"
End Sub
'##ModelId=3A0F61DD03AC
Sub ResetControls()
txtEntry(2).Text = ""
txtEntry(3).Text = ""
txtCallNote.Text = ""
txtMinutes.Text = "Enter Minutes Here"
lblCaseID.Caption = "0"

lstItem(2).ListIndex = -1
lstItem(3).ListIndex = -1

End Sub

Private Sub Form_Activate()
    If Not cmdForceTVWLoad.Enabled Then
        cmdForceTVWLoad.Enabled = True
        LoadAllCustomers clCOMPANY
    End If
End Sub

Private Sub Form_GotFocus()
Screen.MousePointer = vbDefault
End Sub

'##ModelId=3A0F61DE0032
Private Sub Form_Initialize()

InitForm
InitLists

End Sub

Private Sub Form_Load()
'LoadAllCustomers clCOMPANY
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set Employees = Nothing
Set Entries = Nothing
If Not rsGeneric Is Nothing Then Set rsGeneric = Nothing
If Not Employees Is Nothing Then Set Employees = Nothing
If Not colLinks Is Nothing Then Set colLinks = Nothing
If Not Customers Is Nothing Then Set Customers = Nothing
If Not Contacts Is Nothing Then Set Contacts = Nothing
If Not Products Is Nothing Then Set Products = Nothing
If Not CallCodes Is Nothing Then Set CallCodes = Nothing
If Not frmSplash Is Nothing Then Set frmSplash = Nothing
If Not BLServer Is Nothing Then Set BLServer = Nothing
End
End Sub

Private Sub Form_Resize()
Dim xScreenRes As Long
Dim yScreenRes As Long
xScreenRes = GetSystemMetrics(0)
yScreenRes = GetSystemMetrics(1)
End Sub

Private Sub lblCaseID_DblClick()
    txtEnterCaseID.Visible = True
    lblCaseID.Visible = False
    txtEnterCaseID.SetFocus
End Sub
'##ModelId=3A0F61DE0140
Private Sub lstItem_Click(Index As Integer)
Dim vCounter As Variant
Dim iTemp As Long
Dim sText As String
        
'-----------------------------------------------------------------
'If required when clearing the selection
If lstItem(Index).ListIndex <> -1 Then
    iTemp = lstItem(Index).ItemData(lstItem(Index).ListIndex)
    txtEntry(Index).Text = lstItem(Index).Text
    Select Case Index
        Case 0 'Company
            'Clear the contact list box
            lstItem(1).Clear
            'Clear the contact text box
            txtEntry(1).Text = ""
            '-----------------------------------------------------------------
            'load the contacts for the customer
            For Each vCounter In colLinks
                If vCounter.CompanyID = iTemp Then
                    If LoadListBox(lstItem(1), Contacts, vCounter.ContactID) = 0 Then MsgBox "Error"
                End If
            Next vCounter
            '-----------------------------------------------------------------
            GetCompanyHistory iTemp, txtCallHistory
        Case 1 'Contact
        Case 2 'Product
        Case 3 'Call code
    End Select
End If
End Sub

Private Sub tmrBLStatus_Timer()
If BLServer Is Nothing Then
    shBLStatus.FillColor = vbRed
Else
    shBLStatus.FillColor = vbGreen
End If
End Sub

Private Sub tmrMain_Timer()
Static dTime As Double
Static CurrentTime As Date

CurrentTime = Now()

dTime = CurrentTime - StartTime
txtMinutes.Text = Format((dTime), "hh:mm:ss")
End Sub

Private Sub tvwCustomers_Click()
'-----------------------------------------------------------------
'If is required when clearing the selection
'Static sCurrentParent As String

On Error GoTo tvwCustomersErrorHandler
If Not cboEmployees.ListIndex = 0 Then
cmdEditContact.Enabled = True
cmdNewContact.Enabled = True
End If

With tvwCustomers.SelectedItem
    If .Parent Is Nothing Then
        'This is a CUSTOMER (parent)
        If sCurrentParent <> .Key Then
            LoadOneCustomer clCONTACT3, Right(.Key, Len(.Key) - 1)
            GetCompanyHistory Right(.Key, Len(.Key) - 1), txtCallHistory
            lblCaseID.Caption = "0"
        End If
        .Expanded = True
        txtEntry(0).Text = .Text
        txtEntry(1).Text = ""
    Else
        'It is a CONTACT (child)
        txtEntry(0).Text = .Parent.Text
        txtEntry(1).Text = .Text
        If sCurrentParent <> .Parent.Key Then
            GetCompanyHistory Right(.Parent.Key, Len(.Parent.Key) - 1), txtCallHistory
            lblCaseID.Caption = "0"
        End If
    End If
    If InStr(1, .Key, "p") Then sCurrentParent = .Key Else sCurrentParent = .Parent.Key
'    If Left(.Key, 1) = "p" Then sCurrentParent = .Key Else sCurrentParent = .Parent.Key
End With
Exit Sub
tvwCustomersErrorHandler:
Select Case Err.Number
    Case 3021
        MsgBox tvwCustomers.SelectedItem.Text & " has no contacts, add a contact before proceeding.", vbInformation, "CL-FMain::LoadCustomer"
    Case Else
        MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadCustomer"
End Select
End Sub

Private Sub tvwCustomers_DragDrop(Source As Control, X As Single, Y As Single)
Dim lEmplID As Long

If Not tvwCustomers.DragIcon = imgTreeViewIcons.ListImages(clNODROP).Picture Then
    lEmplID = cboEmployees.ItemData(cboEmployees.ListIndex)
'=====================================
'Need to create a EditCompanyLink function to change companyid
'    lEmplID = BLServer.AddCompanyLink(iNewCompanyID, iSelectedContactID, lEmplID)
'    ReInit -1
End If
tvwCustomers.Drag vbEndDrag
tvwCustomers.DropHighlight = Nothing
End Sub

Private Sub tvwCustomers_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'Set highlight to the DRAGGED OVER node
If Not cboEmployees.ItemData(cboEmployees.ListIndex) = 0 Then
    With tvwCustomers
        If Not .HitTest(X, Y) Is Nothing Then 'we're over a node
            Set .DropHighlight = .HitTest(X, Y)
            'Test whether it's a parent
            'Get contact's ID
            iNewCompanyID = CLng(Right(tvwCustomers.DropHighlight.Key, Len(tvwCustomers.DropHighlight.Key) - 1))
            If (Left(.DropHighlight.Key, 1) = "p") And (iNewCompanyID <> iSelectedCompanyID) Then
            'we're over a different parent node
                Source.DragIcon = imgTreeViewIcons.ListImages(Source.SelectedItem.Image + 21).Picture
            Else
                Source.DragIcon = imgTreeViewIcons.ListImages(clNODROP).Picture
            End If
        End If
    End With
End If
End Sub

Private Sub tvwCustomers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lEmplID As Long
    With tvwCustomers
        If Not .HitTest(X, Y) Is Nothing Then
            .SelectedItem = .HitTest(X, Y)
            If Not cboEmployees.ItemData(cboEmployees.ListIndex) = 0 Then
                If Not (.SelectedItem.Parent Is Nothing) Then
                    'Get contact's ID
                    iSelectedContactID = CLng(Right(.SelectedItem.Key, Len(.SelectedItem.Key) - 1))
                    iSelectedCompanyID = CLng(Right(.SelectedItem.Parent.Key, Len(.SelectedItem.Parent.Key) - 1))
                    fDrag = True
                End If
            End If
        End If
    End With
End Sub

Private Sub tvwCustomers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton And fDrag And (cboEmployees.ItemData(cboEmployees.ListIndex) <> 0) Then
    tvwCustomers.Drag vbBeginDrag
Else
    fDrag = False
End If
End Sub

Private Sub tvwCustomers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not tvwCustomers.HitTest(X, Y) Is Nothing Then tvwCustomers.SelectedItem = tvwCustomers.HitTest(X, Y)
End Sub

Private Sub txtEnterCaseID_DblClick()
    txtEnterCaseID.SelStart = 0
    txtEnterCaseID.SelLength = Len(txtEnterCaseID.Text)
End Sub

Private Sub txtEnterCaseID_GotFocus()
    txtEnterCaseID.SelStart = 0
    txtEnterCaseID.SelLength = Len(txtEnterCaseID.Text)
End Sub

Private Sub txtEnterCaseID_KeyPress(KeyAscii As Integer)
Dim lCompanyKey As Long
On Error GoTo WrongValue

lCompanyKey = 0
Select Case KeyAscii
    Case 13
        KeyAscii = 0
        lblCaseID.Caption = txtEnterCaseID.Text
        lblCaseID.Visible = True
        txtEnterCaseID.Visible = False
        txtEnterCaseID.Text = "0"
        Exit Sub
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
    Case Else
        KeyAscii = 0
End Select

Exit Sub

WrongValue:

Select Case Err
    Case 13
        MsgBox "You must enter the Case Number", , "Error " & Err.Number
    Case Else
        MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::txtEnterCaseID"
End Select

End Sub

Private Sub txtEnterCaseID_LostFocus()
Dim lCompanyKey As Long
Dim sTemp As String
Dim lTemp As Long
    If lblCaseID.Caption <> "" And lblCaseID.Caption <> "Enter ID" Then
        If tvwCustomers.SelectedItem Is Nothing Then
            lblCaseID.Caption = txtEnterCaseID.Text
            lblCaseID.Visible = True
            txtEnterCaseID.Visible = False
            txtEnterCaseID.Text = "0"
            Exit Sub
        End If
        With tvwCustomers.SelectedItem
            If Left(.Key, 1) = "p" Then
                'This is a CUSTOMER (parent)
                lCompanyKey = CLng(Right(.Key, Len(.Key) - 1))
            Else
                'It is a CONTACT (child)
                lCompanyKey = CLng(Right(.Parent.Key, Len(.Parent.Key) - 1))
            End If
        End With
    Else
        lblCaseID.Caption = "0"
    End If 'Is there a call to enter
    
    If lblCaseID.Visible = False Then
        lblCaseID.Caption = txtEnterCaseID.Text
        lblCaseID.Visible = True
        txtEnterCaseID.Visible = False
        txtEnterCaseID.Text = "0"
    End If
    
    lTemp = FindCaseID(lblCaseID.Caption)
    sTemp = "p" & CStr(lTemp)
    
    If lTemp > 0 Then
        With tvwCustomers.Nodes.Item(sTemp)
            .Selected = True
             LoadOneCustomer clCONTACT3, Right(.Key, Len(.Key) - 1)
            .Expanded = True
            GetCompanyHistory lTemp, txtCallHistory, lblCaseID.Caption
        End With
    End If

End Sub

'##ModelId=3A0F61DE0295
Private Sub txtEntry_DblClick(iIndex As Integer)

If cboEmployees.ItemData(cboEmployees.ListIndex) = 0 Then
    MsgBox "Please select an Employee.", vbExclamation, "FMain::txtEntry-DblClick"
Else
    Select Case iIndex
        Case 0 'New Customer
        Case 1 'New Contact
            If lstItem(0).ListIndex <> -1 Then
                FAddItem.CompanyID = lstItem(0).ItemData(lstItem(0).ListIndex)
            Else
                MsgBox "Please select a company.", vbInformation, "FMain::txtEntry-DblClick"
                Exit Sub
            End If
        Case 2 'New Product
                Exit Sub
        Case 3 'New Call Code
    End Select
    
    FAddItem.EmplID = cboEmployees.ItemData(cboEmployees.ListIndex)
    FAddItem.ActionType = iIndex + 1
    FAddItem.EntryText = txtEntry(iIndex).Text
    FAddItem.Show vbModal
    
    If FAddItem.ErrorCode <> 0 Then
        ReInit iIndex
'    Else
'        MsgBox "Error adding item, error = " & FAddItem.ErrorCode, vbCritical, "txtEntry"
    End If
End If
End Sub
'##ModelId=3A0F61DF005B
Private Sub txtEntry_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Call FindListItem(lstItem(Index), txtEntry(Index))
End Sub
'##ModelId=3A0F61DF0354
Private Sub txtMinutes_GotFocus()
    txtMinutes.SelStart = 0
    txtMinutes.SelLength = Len(txtMinutes.Text)
End Sub
'##ModelId=3A0F61DF03E1
Private Sub txtminutes_KeyPress(KeyAscii As Integer)
On Error GoTo WrongValue
Dim Items As Integer
Dim MaxItems As Integer
Dim EntryText As String

If KeyAscii = 13 Then 'Carriage Return?
    KeyAscii = 0
    If txtMinutes.Text <> "" Then
        Number = Number + 1
        Dim Entry As New CEntry
        Set Entry = New CEntry
        
        Entry.EmplID = EmplID
        Entry.ETime = txtMinutes.Text
        Entry.EDate = lblDate.Caption
        Entry.Caller = ""
        Entry.Product = ""
        Entry.EIndex = Number
        
'    Add entry to collection
        Entries.Add Item:=Entry, Key:=CStr(Number)
        
        Set Entry = Nothing
        txtMinutes.Text = ""

        For Items = 0 To 4
            Set Entry = Entries.Item(Number - Items)
            EntryText = Entry.EIndex & " : " & Entry.EmplID & " : " & Entry.EDate & " : " & Entry.ETime
            If Entries.Count >= Items + 1 Then lblEntryData(Items).Caption = EntryText
            If Entries.Count - 1 <= Items Then Exit For
        Next
        
        If Entries.Count > 5 Then MaxItems = 4 Else MaxItems = Entries.Count - 1
        
        For Items = 0 To MaxItems
            cmdDelCall(Items).Enabled = True
            If Entries.Count - 1 <= Items Then Exit For
        Next
    
    End If 'Is there a call to enter
vscCalls.Max = Entries.Count
vscCalls.Value = Entries.Count
lblTotalCalls.Caption = "Total Calls Entered = " & CStr(Entries.Count)

End If 'enter key

Exit Sub

WrongValue:

Select Case Err
    Case 13
        MsgBox "You must enter the Number of minutes to complete the call", , "Error " & Err.Number
    Case Else
        MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::txtMinutes"
End Select

End Sub

'##ModelId=3A0F61E00157
Private Sub vsbChangeDate_Change()
    lblDate.Caption = Format(DateValue(Now()) - vsbChangeDate.Value, "medium date")
End Sub
'##ModelId=3A0F61E001ED
Private Sub vsccalls_Change()
Dim MaxItems As Integer
Dim Items As Integer
Dim EntryText As String

    If Entries.Count < 5 Then vscCalls.Min = Entries.Count
    If Entries.Count >= 5 Then vscCalls.Min = 5
    If Entries.Count = 0 Then Exit Sub
    If Entries.Count >= 5 Then MaxItems = 4 Else MaxItems = Entries.Count - 1

    For Items = 0 To MaxItems
        Set Entry = Entries.Item(vscCalls.Value - Items)
        EntryText = Entry.EIndex & " : " & Entry.EmplID & " : " & Entry.EDate & " : " & Entry.ETime
        If Entries.Count >= Items + 1 Then lblEntryData(Items).Caption = EntryText
        If Entries.Count - 1 <= Items Then Exit For
    Next
End Sub
Private Sub LoadLists(iIndex As Integer, ByRef oData As Object)
On Error GoTo LoadListsErrorHandler

If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    shBLStatus.FillColor = vbGreen
    If BLServer Is Nothing Then
        shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server, contact administrator." & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadLists"
        Exit Sub
    End If
End If

Set rsGeneric = BLServer.GetLbData(iIndex)
If rsGeneric Is Nothing Then
    MsgBox "Recordset not created" & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadLists"
    Exit Sub
End If

rsGeneric.MoveLast
frmSplash.pgbLoadTVW.Value = frmSplash.pgbLoadTVW.Min
frmSplash.pgbLoadTVW.Max = rsGeneric.AbsolutePosition

rsGeneric.MoveFirst
Do While Not rsGeneric.EOF
    Select Case iIndex
        Case clLINK
            If Not (IsNull(rsGeneric!ID) Or IsNull(rsGeneric!CompanyID) Or IsNull(rsGeneric!ContactID)) Then
                If oData.Add(rsGeneric!ID, rsGeneric!CompanyID, rsGeneric!ContactID) Is Nothing Then MsgBox "Error adding Customer", , "CL-FMain::LoadLists"
            End If
        Case Else
            If Not (IsNull(rsGeneric!ID) Or IsNull(rsGeneric!sName)) Then
                If oData.Add(rsGeneric!sName, rsGeneric!ID) Is Nothing Then MsgBox "Error adding Customer", , "FMain-InitLists"
            End If
    End Select
    frmSplash.pgbLoadTVW.Value = rsGeneric.AbsolutePosition
    rsGeneric.MoveNext
Loop

Set rsGeneric = Nothing
Set BLServer = Nothing

Exit Sub

LoadListsErrorHandler:
MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadLists"
Set rsGeneric = Nothing
End Sub
Sub LoadCustomer(iIndex As Integer)
Dim NewNode As Node
Dim rsLinks As Recordset
Dim sParentID As String
Dim sNodeName As String
Dim iCompanyImage As Integer
Dim rsGeneric As ADODB.Recordset

On Error GoTo LoadCustomerErrorHandler

If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    shBLStatus.FillColor = vbGreen
    If BLServer Is Nothing Then
        shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server" & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadLists"
        Exit Sub
    End If
End If

Set rsGeneric = BLServer.GetLbData(iIndex)
If rsGeneric Is Nothing Then
    MsgBox "Recordset not created"
    Exit Sub
End If

rsGeneric.MoveFirst
Select Case iIndex
    Case clCOMPANY
        Do While Not rsGeneric.EOF
            If Not (IsNull(rsGeneric!ID) Or IsNull(rsGeneric!sName)) Then
                sParentID = "p" & CStr(rsGeneric!ID)
                sNodeName = rsGeneric!sName ' & " - p" & CStr(rsGeneric!ID)
                Select Case rsGeneric!cType
                    Case "D"
                        iCompanyImage = 1
                    Case "E"
                        iCompanyImage = 2
                    Case "F"
                        iCompanyImage = 3
                    Case "I"
                        iCompanyImage = 4
                    Case "O"
                        iCompanyImage = 5
                    Case "R"
                        iCompanyImage = 6
                    Case "S"
                        iCompanyImage = 7
                    Case "U"
                        iCompanyImage = 8
                    Case Else
                        iCompanyImage = 8
                End Select
                Set NewNode = tvwCustomers.Nodes.Add(, , sParentID, sNodeName, iCompanyImage, iCompanyImage + 20)
            End If
            rsGeneric.MoveNext
        Loop
    Case clCONTACT
        Set rsLinks = BLServer.GetLbData(clLINK)
        rsLinks.MoveFirst
        Do While Not rsLinks.EOF
            If Not (IsNull(rsLinks!ID) Or IsNull(rsLinks!CompanyID) Or IsNull(rsLinks!ContactID)) Then
                rsGeneric.Find "ID = '" & rsLinks!ContactID & "'", 1
                If rsGeneric.EOF = False Then
                    If Not (IsNull(rsGeneric!ID) Or IsNull(rsGeneric!sName)) Then
                        If InStr(1, rsGeneric!Training, "A") <> 0 Or _
                            InStr(1, rsGeneric!Training, "D") <> 0 Or _
                            InStr(1, rsGeneric!Training, "S") <> 0 Then
                            iCompanyImage = 9 + CInt(rsGeneric!Skill) - 1
                        End If
                        If InStr(1, rsGeneric!Training, "L") <> 0 Or _
                            InStr(1, rsGeneric!Training, "R") <> 0 Then
                            iCompanyImage = 12 + CInt(rsGeneric!Skill) - 1
                        End If
                        If InStr(1, rsGeneric!Training, "C") <> 0 Then
                            iCompanyImage = 15 + CInt(rsGeneric!Skill) - 1
                        End If
                        If InStr(1, rsGeneric!Training, "N") Then
                            iCompanyImage = 18 + CInt(rsGeneric!Skill) - 1
                        End If
                        If iCompanyImage = 0 Then iCompanyImage = 18
                        Set NewNode = tvwCustomers.Nodes.Add("p" & CStr(rsLinks!CompanyID), tvwChild, "c" & CStr(rsLinks!ContactID), CStr(rsGeneric!sName), iCompanyImage)
                    End If
                End If
            End If
            rsGeneric.MoveFirst
            rsLinks.MoveNext
        Loop
        Set rsLinks = Nothing
    Case clLINK
        If Not (IsNull(rsGeneric!ID) _
            Or IsNull(rsGeneric!CompanyID) _
            Or IsNull(rsGeneric!ContactID)) Then
            Set NewNode = tvwCustomers.Nodes.Add(, , "p" & CStr(rsGeneric!ID), "Company " & rsGeneric!CompanyID, 1)
        End If
    Case Else '
End Select
iCompanyImage = -1

Set NewNode = Nothing
Set rsGeneric = Nothing
Set rsLinks = Nothing
Set BLServer = Nothing

Exit Sub

LoadCustomerErrorHandler:
MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadCustomer"
Set rsGeneric = Nothing
End Sub

Sub LoadCustomer2(iIndex As Integer)
Dim NewNode As Node
Dim sParentID As String
Dim sChildID As String
Dim sNodeName As String
Dim iCompanyImage As Integer
Dim rsGeneric As ADODB.Recordset

On Error GoTo LoadCustomer2ErrorHandler

'Make sure a server exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    shBLStatus.FillColor = vbGreen
    If BLServer Is Nothing Then
        shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server" & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadLists"
        Exit Sub
    End If
End If

'Get the customer data from the server
Set rsGeneric = BLServer.GetLbData(iIndex)
If rsGeneric Is Nothing Then
    MsgBox "Recordset not created"
    Exit Sub
End If

'Fill the treeview
rsGeneric.MoveFirst
Do While Not rsGeneric.EOF
    sParentID = CStr(rsGeneric!ParentID)
    sChildID = CStr(rsGeneric!ChildID)
    sNodeName = rsGeneric!sName
    
    Select Case iIndex
        Case clCOMPANY2
            Set NewNode = tvwCustomers.Nodes.Add(, , sChildID, sNodeName, CInt(rsGeneric!cType), CInt(rsGeneric!cType) + 20)
        Case clCONTACT2
            Set NewNode = tvwCustomers.Nodes.Add(sParentID, tvwChild, sChildID, sNodeName, CInt(rsGeneric!cType), CInt(rsGeneric!cType))
    End Select
    rsGeneric.MoveNext
Loop

iCompanyImage = -1

Set NewNode = Nothing
Set rsGeneric = Nothing
Set BLServer = Nothing

Exit Sub

LoadCustomer2ErrorHandler:
MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadCustomer2"
Set rsGeneric = Nothing
'Resume Next
End Sub

Sub LoadAllCustomers(iIndex As Integer)
Dim rsGeneric As ADODB.Recordset
Dim NewNode As Node
Dim sParentID As String
Dim sChildID As String
Dim sNodeName As String
Dim iCounter As Long
Dim iType As Integer
Dim iCompanyImage As Integer

On Error GoTo LoadAllCustomersErrorHandler

frmSplash.MousePointer = vbHourglass
iCounter = 0

'Make sure a server exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    shBLStatus.FillColor = vbGreen
    If BLServer Is Nothing Then
        shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server" & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadLists"
        Exit Sub
    End If
End If

'Get the customer data from the server
Set rsGeneric = BLServer.GetLbData(iIndex)
If rsGeneric Is Nothing Then
    MsgBox "Recordset not created"
    Exit Sub
End If
Set BLServer = Nothing

'Fill the treeview
If frmSplash.Visible = True Then
    rsGeneric.MoveLast
    frmSplash.pgbLoadTVW.Value = frmSplash.pgbLoadTVW.Min
    frmSplash.pgbLoadTVW.Max = rsGeneric.AbsolutePosition
End If

rsGeneric.MoveFirst

tvwCustomers.Sorted = False

If fQuery.Visible = False Then
    If frmSplash.Visible = False Then
        tvwCustomers.Nodes.Clear
        Do While Not rsGeneric.EOF
            sParentID = (rsGeneric!ParentID)
            sChildID = (rsGeneric!ChildID)
            sNodeName = rsGeneric!sName
            iType = CInt(rsGeneric!cType)
            With tvwCustomers.Nodes
                Select Case iIndex
                    Case clCOMPANY
                        Set NewNode = .Add(, , sChildID, sNodeName, iType, iType + 20)
                        NewNode.Expanded = False
                        sbMain.Panels(3).Text = "Loading Companies ..." & sParentID
'                        frmSplash.Label1.Caption = "Loading Companies ..." & sParentID
'                        frmSplash.txtJWVTest.Text = frmSplash.txtJWVTest.Text & sParentID & " :: " & sChildID & " :: " & sNodeName & " :: " & iType & vbCrLf
                    Case clCONTACT
                        Set NewNode = .Add(sParentID, tvwChild, sChildID, sNodeName, iType, iType + 21)
                        NewNode.Expanded = False
                        sbMain.Panels(3).Text = "Loading Contact ..." & sChildID & " to " & sParentID
'                        frmSplash.Label1.Caption = "Loading Contacts ..." & sChildID
                End Select
            End With
'            frmSplash.pgbLoadTVW.Value = rsGeneric.AbsolutePosition
'            frmSplash.Refresh
            rsGeneric.MoveNext
            DoEvents
        Loop
        sbMain.Panels(3).Text = ""
    Else 'on the splash screen
        Do While Not rsGeneric.EOF
            sParentID = (rsGeneric!ParentID)
            sChildID = (rsGeneric!ChildID)
            sNodeName = rsGeneric!sName
            iType = CInt(rsGeneric!cType)
            With tvwCustomers.Nodes
                Select Case iIndex
                    Case clCOMPANY
                        Set NewNode = .Add(, , sChildID, sNodeName, iType, iType + 20)
                        NewNode.Expanded = False
                        frmSplash.Label1.Caption = "Loading Companies ..." & sParentID
                        frmSplash.txtJWVTest.Text = frmSplash.txtJWVTest.Text & sParentID & " :: " & sChildID & " :: " & sNodeName & " :: " & iType & vbCrLf
'                        sbMain.Panels(3).Text = "Loading Companies ..." & sParentID
                    Case clCONTACT
                        Set NewNode = .Add(sParentID, tvwChild, sChildID, sNodeName, iType, iType + 21)
                        NewNode.Expanded = False
                        frmSplash.Label1.Caption = "Loading Contacts ..." & sChildID
'                        sbMain.Panels(3).Text = "Loading Contact ..." & sChildID & " to " & sParentID
                End Select
            End With
            frmSplash.pgbLoadTVW.Value = rsGeneric.AbsolutePosition
            rsGeneric.MoveNext
            frmSplash.Refresh
            DoEvents
        Loop
    End If
Else
    fQuery.txtCallHistory.Text = ""
    iCounter = 0
    Do While Not rsGeneric.EOF
        fQuery.txtCallHistory.Text = fQuery.txtCallHistory.Text & vbCrLf & _
            rsGeneric!sName & " (" & rsGeneric!ParentID & ")"
        rsGeneric.MoveNext
        DoEvents
        iCounter = iCounter + 1
    Loop
    fQuery.txtCallHistory.Text = fQuery.txtCallHistory.Text & vbCrLf & vbCrLf & _
    "Number of entries = " & iCounter
End If

'frmSplash.Label1.Caption = "Loading Companies ..." & "DONE"

iCompanyImage = -1
'sbMain.Panels(3).Text = ""
'frmSplash.MousePointer = vbDefault

Set NewNode = Nothing
Set rsGeneric = Nothing

tvwCustomers.Sorted = True

Exit Sub

LoadAllCustomersErrorHandler:

    MsgBox "Actual Node Count: " & SendMessage(tvwCustomers.hwnd, _
           TVM_GETCOUNT, 0, ByVal 0)
    MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadAllCustomers"
Set rsGeneric = Nothing
'Resume Next
frmSplash.MousePointer = vbDefault
End Sub

Sub LoadOneCustomer(iIndex As Integer, Optional Filter As Long)
Dim NewNode As Node
Dim sParentID As String
Dim sChildID As String
Dim sNodeName As String
Dim iCompanyImage As Integer
'Dim rsGeneric As ADODB.Recordset
Dim iCounter As Long
Dim iType As Integer

On Error GoTo LoadOneCustomerErrorHandler
Set rsGeneric = Nothing
sParentID = ""
sChildID = ""
sNodeName = ""

Screen.MousePointer = vbHourglass

'Make sure a server exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    shBLStatus.FillColor = vbGreen
    If BLServer Is Nothing Then
        shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server, notify administrator.", vbCritical, "CL-FMain::LoadLists"
        Exit Sub
    End If
End If

'Get the customer data from the server
Set rsGeneric = BLServer.GetLbData(iIndex, Filter)
If rsGeneric Is Nothing Then
    MsgBox "Unable to get Contacts, notify administrator.", vbCritical, "CL-Main::LoadOneCustomer"
    Exit Sub
End If
Set rsGeneric.ActiveConnection = Nothing
Set BLServer = Nothing

'Fill the treeview
rsGeneric.MoveFirst

sParentID = CStr(rsGeneric!ParentID)
sChildID = CStr(rsGeneric!ChildID)
sNodeName = rsGeneric!sName
iType = CInt(rsGeneric!cType)

With tvwCustomers.Nodes
    Select Case iIndex
        'Add a new company node
        Case clCOMPANY
            Set NewNode = .Add(, , sChildID, sNodeName, iType, iType + 20)
            NewNode.Expanded = False
            sbMain.Panels(3).Text = "Loading Companies ..." & sParentID
        'Edit an existing company node
        Case clCOMPANY2
            With .Item(sParentID)
                .Text = sNodeName
                .Image = iType
                .SelectedImage = iType + 20
            End With
        'Add a new contact node
        Case clCONTACT
            Set NewNode = .Add(sParentID, tvwChild, sChildID, sNodeName, iType, iType + 21)
            NewNode.Expanded = False
            sbMain.Panels(3).Text = "Loading Contact ..." & sChildID & " to " & sParentID
        'Edit an existing contact node
        Case clCONTACT2
            With .Item(sChildID)
                .Text = sNodeName
                .Image = iType
                .SelectedImage = CInt(rsGeneric!cType) + 21
                tvwCustomers.Refresh
            End With
        'Refresh the list of contacts (delete them then reload)
        Case clCONTACT3
            If tvwCustomers.SelectedItem.Children > 0 Then
                For iCounter = 1 To tvwCustomers.SelectedItem.Children
                    .Remove (tvwCustomers.SelectedItem.Child.FirstSibling.Index)
                Next iCounter
            End If
            Do While Not rsGeneric.EOF
                Set NewNode = .Add(CStr(rsGeneric!ParentID), tvwChild, CStr(rsGeneric!ChildID), rsGeneric!sName, CInt(rsGeneric!cType), CInt(rsGeneric!cType) + 21)
                rsGeneric.MoveNext
            Loop
        'Add a new company node
        Case clCOMPANYUPDATE
    End Select
End With

DoEvents

iCompanyImage = -1

Set NewNode = Nothing
Set rsGeneric = Nothing
'Set BLServer = Nothing
Screen.MousePointer = vbDefault

Exit Sub

LoadOneCustomerErrorHandler:
Set BLServer = Nothing
Set NewNode = Nothing
Set rsGeneric = Nothing

Select Case Err.Number
    Case 3021
        If Not tvwCustomers.SelectedItem Is Nothing Then
            MsgBox "(" & tvwCustomers.SelectedItem.Text & ") has no contacts, add a contact before proceeding.", vbInformation, "CL-FMain::LoadOneCustomer"
        Else
'            MsgBox "No customer selected.", vbInformation, "CL-FMain::LoadOneCustomer"
        End If
    Case 35601, 35602
    Case Else
        MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadOneCustomer"
End Select
'MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadOneCustomer"
Set rsGeneric = Nothing
Screen.MousePointer = vbDefault

End Sub

Private Sub BLServer_OnUpdateDone()
'MsgBox "BLServer event OnUpdateDone fired"
'Debug.Print "BLServer event OnUpdateDone"
'ReInit Me.LastAction
'
''Clear last action
'Me.LastAction = clDEFAULTLASTACTION
'Me.LastCompany = clDEFAULTLASTCOMPANY
'Me.LastContact = clDEFAULTLASTCONTACT
End Sub

Property Get LastAction() As Long
LastAction = iLastAction
End Property

Property Let LastAction(ByVal vNewValue As Long)
iLastAction = vNewValue
End Property

Property Get LastCompany() As Long
LastCompany = iLastCompany
End Property

Property Let LastCompany(ByVal vNewValue As Long)
iLastCompany = vNewValue
End Property

Property Get LastContact() As Long
LastContact = iLastContact
End Property

Property Let LastContact(ByVal vNewValue As Long)
iLastContact = vNewValue
End Property

Private Function FindCaseID(sCaseID As String) As Long
On Error GoTo FindCaseIDErrorHandler
FindCaseID = 0

'Make sure a server exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    shBLStatus.FillColor = vbGreen
    If BLServer Is Nothing Then
        shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server" & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadLists"
        Exit Function
    End If
End If

'Get the customer data from the server
Set rsGeneric = BLServer.GetLbData(clCASEID, CLng(sCaseID))
If rsGeneric Is Nothing Then
    MsgBox "Recordset not created"
    Exit Function
Else
    
    If (rsGeneric.BOF And rsGeneric.EOF) Then
        txtCallHistory.Text = vbCrLf & vbCrLf & vbCrLf
        txtCallHistory.Text = txtCallHistory.Text & vbCrLf & vbCrLf & vbCrLf
        txtCallHistory.Text = txtCallHistory.Text & vbCrLf & vbCrLf & vbCrLf & "     ============================" & vbCrLf
        txtCallHistory.Text = txtCallHistory.Text & "            Invalid Case Id      " & vbCrLf
        txtCallHistory.Text = txtCallHistory.Text & "     ============================" & vbCrLf
        sCurrentParent = 0
    Else
        FindCaseID = rsGeneric!CompanyID
        sCurrentParent = "p" & rsGeneric!CompanyID
    End If
End If

Set rsGeneric = Nothing
Set BLServer = Nothing

Exit Function

FindCaseIDErrorHandler:
MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::FindCaseID"

Set rsGeneric = Nothing
Set BLServer = Nothing

End Function
'Private Sub gExit(Optional Cancel As Integer, _
'                  Optional UserReqst As Boolean = True)
'Dim i As Integer
'
'On Error Resume Next
'For i = Forms.Count - 1 To 0 Step -1
'    Unload Forms(i)
'    Set Forms(i) = Nothing
'Next i
'End
'End Sub
''========================
'Private Sub gExit2(ByRef frm As Form)
'Dim i  As Integer
'Dim RmvalFlag As Boolean
'
'On Error Resume Next
'Screen.MousePointer = vbHourglass
'
''code to do pre-close stuff
'
''==== TERMINATES/KILLS
''Loop round all forms to KILL them
'For i = Forms.Count - 1 To 0 Step -1
'    If Forms(i).Name <> Main.Name Then
'        Unload Forms(i)
'    End If
'Next i
'
''any additional codcode for
'
''Kill the MAIN reference
'Set Main = Nothing
'Set frmMain = Nothing
'Screen.MousePointer = vbDefault
'End
'End Sub

