VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FMain2 
   Caption         =   "Call Tracker 2"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12045
   FillColor       =   &H80000013&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "FProdSupLoading2.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8715
   ScaleWidth      =   12045
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
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "Enter ID"
      Top             =   585
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer tmrBLStatus 
      Interval        =   1000
      Left            =   8640
      Top             =   5520
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
      TabIndex        =   32
      TabStop         =   0   'False
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
      Left            =   1800
      TabIndex        =   6
      Text            =   "Enter Minutes Here"
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query"
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
      TabIndex        =   30
      TabStop         =   0   'False
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
      Left            =   120
      MaskColor       =   &H00FFFF00&
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdNewCustomer 
      Caption         =   "Add New Company"
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
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdNewContact 
      Caption         =   "Add New   Contact"
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
      TabIndex        =   27
      TabStop         =   0   'False
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "   Add   New Call"
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
      TabIndex        =   11
      Top             =   1140
      Width           =   1575
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9120
      Top             =   5520
   End
   Begin VB.CommandButton cmdEditCall 
      Caption         =   "   Reset         Form"
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1575
   End
   Begin ComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   8460
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "10/14/2002"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "11:03 PM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   13053
            MinWidth        =   13053
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   0
            Text            =   "FMain"
            TextSave        =   "FMain"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmCall 
      Height          =   3475
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   10215
      Begin VB.CommandButton cmdGetOpenCalls 
         Caption         =   "Get Open Calls"
         Height          =   255
         Left            =   2280
         TabIndex        =   45
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optCallCode 
         Caption         =   "Start-up Help"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtCallNote 
         Height          =   1395
         Index           =   1
         Left            =   3960
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1980
         Width           =   6135
      End
      Begin VB.CheckBox chkCallComplete 
         Caption         =   "Call Complete"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   420
         Width           =   3735
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   3735
      End
      Begin VB.OptionButton optCallCode 
         Caption         =   "Application"
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
      End
      Begin VB.OptionButton optCallCode 
         Caption         =   "Troubleshooting"
         Height          =   255
         Index           =   11
         Left            =   2280
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2640
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optCallCode 
         Caption         =   "General Info"
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtCallNote 
         Height          =   1395
         Index           =   0
         Left            =   3960
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   6135
      End
      Begin VB.ListBox lstItem 
         Height          =   1620
         Index           =   2
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblCallNote 
         Caption         =   "Answers"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   1
         Left            =   3960
         TabIndex        =   42
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblContact 
         Caption         =   "Contact Name"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblCustomer 
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   175
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblCallNote 
         Caption         =   "Questions"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   3960
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblProduct 
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   120
         TabIndex        =   17
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label lblCallCode 
         Caption         =   "Call Code"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   2280
         TabIndex        =   16
         Top             =   1260
         Width           =   1575
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
      Left            =   10380
      TabIndex        =   14
      Top             =   7560
      Width           =   1575
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
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   4335
   End
   Begin VB.VScrollBar vsbChangeDate 
      Height          =   495
      Left            =   5640
      Max             =   0
      Min             =   30
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   110
      Width           =   255
   End
   Begin VB.Frame frmCustomer 
      Height          =   4335
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   10215
      Begin MSFlexGridLib.MSFlexGrid mfgHistory 
         DragIcon        =   "FProdSupLoading2.frx":0442
         Height          =   3795
         Left            =   3960
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   420
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   6694
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   50
         WordWrap        =   -1  'True
         HighLight       =   2
         GridLinesFixed  =   1
         MergeCells      =   2
         AllowUserResizing=   3
         FormatString    =   "Date|Note|Case|Tech|Product"
      End
      Begin VB.ListBox lstItem 
         Height          =   645
         Index           =   3
         Left            =   4080
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   3
         Left            =   4080
         TabIndex        =   40
         Top             =   2880
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   2
         Left            =   4080
         TabIndex        =   39
         Top             =   2520
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtCallHistory 
         Height          =   2775
         Left            =   6000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   33
         Text            =   "FProdSupLoading2.frx":0884
         Top             =   660
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.ListBox lstItem 
         Height          =   645
         Index           =   0
         Left            =   4080
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ListBox lstItem 
         Height          =   1230
         Index           =   1
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   3000
         Width           =   3735
      End
      Begin ComctlLib.TreeView tvwCustomers 
         DragIcon        =   "FProdSupLoading2.frx":088A
         Height          =   2535
         Left            =   120
         TabIndex        =   24
         Top             =   420
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   4471
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imgTreeViewIcons"
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Companies and Contacts"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   175
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblTVKey 
         Height          =   255
         Left            =   7440
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Call History"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   175
         Left            =   3960
         TabIndex        =   20
         Top             =   240
         Width           =   2295
      End
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
      Left            =   9600
      Top             =   5400
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
            Picture         =   "FProdSupLoading2.frx":0CCC
            Key             =   "D"
            Object.Tag             =   "D"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":0FE6
            Key             =   "E"
            Object.Tag             =   "E"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":1300
            Key             =   "F"
            Object.Tag             =   "F"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":161A
            Key             =   "I"
            Object.Tag             =   "I"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":1934
            Key             =   "O"
            Object.Tag             =   "O"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":1C4E
            Key             =   "R"
            Object.Tag             =   "R"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":1F68
            Key             =   "S"
            Object.Tag             =   "S"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":2282
            Key             =   "U"
            Object.Tag             =   "U"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":259C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":28B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":2BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":2EEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":3204
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":351E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":3838
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":3B52
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":3E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":4186
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":44A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":47BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":4AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":4DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":5108
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":5422
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":573C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":5A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":5D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":608A
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":63A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":66BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":69D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":6CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":700C
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":7326
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":7640
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":795A
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":7C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":7F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":82A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":85C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":88DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FProdSupLoading2.frx":8BF6
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
      TabIndex        =   22
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
      TabIndex        =   21
      ToolTipText     =   "Double-click to enter Case ID"
      Top             =   585
      Width           =   1575
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
      Left            =   3960
      TabIndex        =   7
      Top             =   175
      Width           =   1575
   End
End
Attribute VB_Name = "FMain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC85C0112"
Option Explicit

Dim Employees As CEmployees
Dim Customers As CCustomers
Dim Contacts As CContacts
Dim Products As CProducts
Dim CallCodes As CCallCodes
Dim Entries As Collection
Dim Entry As CEntry
Dim colLinks As CLinks
Dim Calls As CCalls

Dim Number As Integer
Dim EmplID As Long
Dim ErrorCondition As Integer
Dim rsGeneric As ADODB.Recordset

Const clCOMPANYBYDATE As Integer = 0
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
Const clCONTACTBYID As Integer = 12
Const clCOMPANYBYID As Integer = 13
Const clCONTACT3 As Integer = 14
Const clCASEID As Integer = 15
Const clCOMPANY4 As Integer = 16
Const clNEWCALLS As Integer = 17
Const clNEWCOMPANIES As Integer = 18
Const clNEWCOMPANIESNOCALLS As Integer = 19
Const clREFRESHCOMPANIES As Integer = 20
Const clREFRESHCONTACTS As Integer = 21

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

Dim WithEvents BLServer As CallTrackerBLServer.BLServer
Attribute BLServer.VB_VarHelpID = -1

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const TVM_GETCOUNT = &H1105&
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                        (ByVal hwnd As Long, _
                        ByVal wMsg As Long, ByVal _
                        wParam As Long, _
                        lParam As Any) As Long
Function InitLists() As Integer
On Error GoTo InitListsErrorHandler

Set Employees = New CEmployees
Set Products = New CProducts

'Connect to the BL Server
frmSplash.Label1.Caption = "Initializing BL Server ..."

'Fill the local collections
frmSplash.Label1.Caption = "Getting Products ..."
LoadLists clPRODUCT, Products
frmSplash.Label1.Caption = "Getting Employees ..."
LoadLists clEMPLOYEE, Employees

'Load the listboxes and comboboxes
frmSplash.Label1.Caption = "Loading Lists ..."
If LoadComboBox(cboEmployees, Employees) = 0 Then MsgBox "error"
If LoadListBox(lstItem(clPRODUCT), Products) = 0 Then MsgBox "Couldn't load listbox", , "BLogic-GetCustomers"

'Load the customer treeview
frmSplash.Label1.Caption = "Loading Companies ..."
dLastCompanyUpdate = Now()
'LoadAllCustomers clCOMPANYBYDATE

'Set Employees = Nothing (required later in program)
Set Products = Nothing

shBLStatus.FillColor = vbRed

InitLists = 0 'Err.Number

Exit Function

InitListsErrorHandler:
MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::InitLists"
End Function

Sub ReInit(Index As Integer)
'============================================================
'This sub reinitializes the collections and list boxes after
'adding a new record for the list (like new customer and contact)
'============================================================

If Index = -1 Then Exit Sub

'Clear the listbox getting new information
Select Case Index
    Case clCOMPANYBYDATE
    'Added a new company and contact
        LoadOneCustomer clCOMPANYBYID, Me.LastCompany 'load parents
        LoadOneCustomer clCONTACT3, Me.LastCompany 'load children
'    Case clCONTACT
'    'Added a new contact
'        LoadOneCustomer clCONTACTBYID, Me.LastContact 'load children
    Case clPRODUCT
    'Added a new product
    Case clCALLCODE
    'Added a new call code (not available in CT2)
'        lstItem(Index).Clear
'        Set CallCodes = New CCallCodes
'        LoadLists clCALLCODE, CallCodes
'        If LoadListBox(lstItem(clCALLCODE), CallCodes) = 0 Then MsgBox "Couldn't load listbox", vbCritical, "FMain::ReInit"
'        GoToItem lstItem(clCALLCODE), FAddItem.CodeID
'        Set CallCodes = Nothing
    Case clCOMPANYUPDATE, clNEWCOMPANIESNOCALLS
    'Changed company data
'        If LoadCurrentCustomers = 0 Then
'            MsgBox "Failed loading current customers", vbCritical, "FMain2::Form_Load"
'        End If
        UpdateOneCustomer clCOMPANYBYID, Me.LastCompany 'load parents
        LoadOneCustomer clCONTACT3, Me.LastCompany 'load children
    Case clNEWCOMPANIESNOCALLS
    'Changed company data
        If LoadCurrentCustomers = 0 Then
            MsgBox "Failed loading current customers", vbCritical, "FMain2::Form_Load"
        End If
'        LoadOneCustomer clCOMPANYBYID, Me.LastCompany 'load parents
    Case clCONTACTUPDATE, clCONTACT
    'Changed contact data
        LoadOneCustomer clCONTACT3, Me.LastCompany 'load children
End Select

'Unload FAddItem
End Sub
Sub InitForm()
On Error GoTo FMainHandler

frmSplash.Label1.Caption = "Initializing ..."

Number = 0

'Initialize Employee combo box
cboEmployees.Clear
'Set Entries = New Collection

sCurrentParent = ""
Me.LastAction = clDEFAULTLASTACTION
Me.LastCompany = clDEFAULTLASTCOMPANY
Me.LastContact = clDEFAULTLASTCONTACT

Load FAddItem 'gets the form into memory for faster adds

'set date
lblDate.Caption = Format(Now(), "medium date")

Exit Sub

FMainHandler:
MsgBox Err.Number & ": " & Err.Description & vbCrLf & Err.Source, vbCritical, "Error InitForm"
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
        Me.LastCompany = CLng(Right(.Key, Len(.Key) - 1))
    End With
End If

'-------------------------------------------------------
'If updating contact, get the contact id
'Get the contact id
If Me.cmdEditContact.Caption = "Edit Contact" Then
    Me.LastContact = iSelectedContactID
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
End If

dLastCompanyUpdate = Now()
End Sub

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
    cmdNewCustomer.Enabled = True
    cmdUpdate.Enabled = True
End If
End Sub

Private Sub chkCallComplete_Click()
Dim sMsgText As String

sMsgText = Me.chkCallComplete.Name & " To Do: " & vbCrLf
sMsgText = sMsgText & "OFF if company/contact not selected" & vbCrLf
sMsgText = sMsgText & "OFF if CompanyID=???" & vbCrLf
sMsgText = sMsgText & "OFF if ContactID=10" & vbCrLf

MsgBox sMsgText
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdEditContact_Click()
'edit company or contact
'check the selected item

cmdNewCustomer.Enabled = False
cmdNewContact.Enabled = False
cmdEditContact.Enabled = False
Screen.MousePointer = vbHourglass

If cmdEditContact.Caption = "Edit Contact" Then
    'edit the contact
    AddItem clCONTACTUPDATE
Else 'edit the company
    AddItem clCOMPANYUPDATE
End If
End Sub

Private Sub cmdForceTVWLoad_Click()
Dim iCounter As Integer

ResetControls
cmdEditContact.Enabled = False
cmdNewContact.Enabled = False
txtCallHistory.Text = ""
sCurrentParent = ""
lstItem(1).Clear
lstItem(1).Refresh
tvwCustomers.Nodes.Clear
tvwCustomers.Refresh
With Me.mfgHistory
    .AddItem "No Data", 1
    For iCounter = 1 To .Rows
        If .Rows > 2 Then
            .RemoveItem 2
        Else
            Exit For
        End If
    Next iCounter
'    .Clear
'    .FormatString = "Date|Note|Case Id|PSG Engr|Contact|Product|Call Type|Duration"
'    .Refresh
End With

'LoadAllCustomers clCOMPANYBYDATE
If LoadCurrentCustomers = 0 Then
    MsgBox "Failed loading current customers", vbCritical, "FMain2::Form_Load"
End If

End Sub

Private Sub cmdGetOpenCalls_Click()
Dim sMsgText As String

sMsgText = Me.cmdGetOpenCalls.Name & " To Do: " & vbCrLf
sMsgText = sMsgText & "Query ID" & vbCrLf
sMsgText = sMsgText & "Get records where EmplID call=EmplID" & vbCrLf
sMsgText = sMsgText & "and CallOpen=TRUE" & vbCrLf

MsgBox sMsgText

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
AddItem (clCOMPANYBYDATE)
'AddItem (clNEWCOMPANIES)
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
Private Sub cmdEditCall_Click()
Dim sText As String
Dim sNote As String
Dim iCustomerID As Long
Dim iContactID As Long
Dim iProductID As Long
Dim iCallCodeID As Long
Dim iEmplID As Long
Dim vCount As Variant
Dim iCallTime As Integer
Dim iStart As Integer

On Error GoTo ErrorHandler

With Me
    If fEditCaseID Then
        ResetControls
'        .cmdEditCall.Caption = "Clear Form"
'        .txtCallNote(0).Text = ""
'        .txtCallNote(1).Text = ""
'        .txtEnterCaseID = "0"
'        .txtEntry(0).Text = ""
'        .txtEntry(1).Text = ""
'        .lblCaseID.Caption = "0"
'        .txtMinutes.Text = "Enter Minutes Here"
'        .lblDate.Caption = Format(Now(), "mmm-dd-yy")
'        .lstItem(clPRODUCT).ListIndex = -1
'
'        fEditCaseID = Not fEditCaseID
            
    Else
    End If
    
    Exit Sub
End With
''===================================
''Check company name was entered
'If Me.txtEntry(0).Text = "" Then
'    MsgBox "Please select company.", vbExclamation, "FMain::cmdUpdate"
'    Me.txtEntry(0).SetFocus
'    Exit Sub
'End If
'
''===================================
''Check that a contact was entered
'If Me.txtEntry(1).Text = "" Then
'    MsgBox "Please select contact.", vbExclamation, "FMain::cmdUpdate"
'    Me.txtEntry(1).SetFocus
'    Exit Sub
'End If
'
''Was the company recognized?
'iStart = InStr(1, Me.txtEntry(0).Text, ":::p")
'
'If Not iStart = 0 Then
'    iCustomerID = Mid(Me.txtEntry(0).Text, iStart + 4, Len(Me.txtEntry(0).Text) - iStart - 3)
'Else
'    iCustomerID = 6 'this really should be the Unknown customer ID
'    sNote = "Company: " & Me.txtEntry(0).Text & vbCrLf
'End If
'
''Was the contact recognized?
'iStart = InStr(1, Me.txtEntry(1).Text, ":::c")
'If Not iStart = 0 Then
'    iContactID = Mid(Me.txtEntry(1).Text, iStart + 4, Len(Me.txtEntry(1).Text) - iStart - 3)
'Else
'    iContactID = 10 'this really should be the Unknown customer ID
'    sNote = sNote & "Contact: " & Me.txtEntry(1).Text & vbCrLf
'End If
'
''===================================
''Check that product was chosen
'
'If lstItem(2).ListIndex = -1 Then
'    MsgBox "Please select product.", vbExclamation, "FMain::cmdUpdate"
'    Exit Sub
'End If
'iProductID = lstItem(2).ItemData(lstItem(2).ListIndex)
'
''===================================
''Choose a call code
'iCallCodeID = GetOption(optCallCode)
'
''===================================
''Find the PSG Engineer
'iEmplID = cboEmployees.ItemData(cboEmployees.ListIndex)
'
''===================================
''Check that call duration was entered
'
'iCallTime = 6
'
'If tmrMain.Enabled = True Then tmrMain.Enabled = False
'
'If Not txtMinutes.Text = "Enter Minutes Here" And Not txtMinutes.Text = "" Then
'    If CLng(CDate(txtMinutes.Text) * 24 * 60) > 480 Then
'        iCallTime = CInt(txtMinutes.Text)
'    Else
'        iCallTime = CInt(CDate(txtMinutes.Text) * 24 * 60)
'    End If
'Else
'    MsgBox "Please enter the call duration.", vbExclamation, "FMain::cmdUpdate"
'    txtMinutes.Text = "Enter Minutes Here"
'    Exit Sub
'End If
'
''===================================
''Check that a note was entered
'
'If txtCallNote(0).Text = "" And txtCallNote(1).Text = "" Then
'    MsgBox "Enter call notes.", vbExclamation, "FMain::cmdUpdate"
'    Exit Sub
'End If
'If Len(Trim(txtCallNote(0).Text)) > 0 Then
'    sNote = sNote & "Questions:" & vbCrLf & txtCallNote(0).Text & vbCrLf
'End If
'If Len(Trim(txtCallNote(1).Text)) > 0 Then
'    sNote = sNote & "Answers:" & vbCrLf & txtCallNote(1).Text
'End If
'
''Connect to the BL Server
'
'If BLServer Is Nothing Then
'    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
'    shBLStatus.FillColor = vbGreen
'    If BLServer Is Nothing Then
'        shBLStatus.FillColor = vbRed
'        MsgBox "Unable to access Server, contact administrator." & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "FMain::cmdUpdate"
'        Exit Sub
'    End If
'End If
'
'If BLServer.AddCall2(iCustomerID, iContactID, iCallCodeID, iProductID, iEmplID, lblDate.Caption, sNote, iCallTime, CLng(lblCaseID.Caption), lCurrentRecordID) = 0 Then
'    MsgBox "Error adding call", vbCritical, "FMain::cmdEditCall"
'    Exit Sub
'End If
'
'txtCallHistory.Text = sText
'GetCompanyHistory iCustomerID, txtCallHistory
'ResetControls
'
'Set BLServer = Nothing

Exit Sub

ErrorHandler:
'    MsgBox "Failed to add call to database : " & Err.Description, vbCritical, "CL-cmdEditCall::CLICK"
End Sub

Private Sub cmdTimer_Click()
Static iTimerOn As Integer

If fEditCaseID Then
    fEditCaseID = False

    With Me
        With .cmdEditCall
            .Enabled = True
            .Caption = "   Reset         Form"
        End With
        .cmdUpdate.Caption = "   Add   New Call"
        .lblDate = Format(Now(), "mmm-dd-yy")
        .txtMinutes.Text = "Enter Minutes Here"
    End With
End If

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
Dim sNote As String
Dim iCustomerID As Long
Dim iContactID As Long
Dim iProductID As Long
Dim iCallCodeID As Long
Dim iEmplID As Long
Dim vCount As Variant
Dim iCallTime As Integer
Dim iStart As Integer

On Error GoTo ErrorHandler

'===================================
'Check company name was entered
    If Me.txtEntry(clCOMPANY).Text = "" Then
        MsgBox "Please select company.", vbExclamation, "FMain::cmdUpdate"
        Me.txtEntry(clCOMPANY).SetFocus
        Exit Sub
    End If

'===================================
'Check that a contact was entered
    If Me.txtEntry(clCONTACT).Text = "" Then
        MsgBox "Please select contact.", vbExclamation, "FMain::cmdUpdate"
        Me.txtEntry(clCONTACT).SetFocus
        Exit Sub
    End If

'Was the company recognized?
    iStart = InStr(1, Me.txtEntry(clCOMPANY).Text, ":::p")
    
    If Not iStart = 0 Then
        iCustomerID = Mid(Me.txtEntry(clCOMPANY).Text, iStart + 4, Len(Me.txtEntry(clCOMPANY).Text) - iStart - 3)
    Else
        iCustomerID = 6 'this really should be the Unknown customer ID
        sNote = "Company: " & Me.txtEntry(clCOMPANY).Text & vbCrLf
    End If

'Was the contact recognized?
    iStart = InStr(1, Me.txtEntry(clCONTACT).Text, ":::c")
    If Not iStart = 0 Then
        iContactID = Mid(Me.txtEntry(clCONTACT).Text, iStart + 4, Len(Me.txtEntry(clCONTACT).Text) - iStart - 3)
    Else
        iContactID = 10 'this really should be the Unknown customer ID
        sNote = sNote & "Contact: " & Me.txtEntry(clCONTACT).Text & vbCrLf
    End If

'===================================
'Check that product was chosen

    If lstItem(2).ListIndex = -1 Then
        MsgBox "Please select product.", vbExclamation, "FMain::cmdUpdate"
        Exit Sub
    End If
    iProductID = lstItem(2).ItemData(lstItem(2).ListIndex)

'===================================
'Choose a call code
    iCallCodeID = GetOption(optCallCode)

'===================================
'Find the PSG Engineer
    iEmplID = cboEmployees.ItemData(cboEmployees.ListIndex)

'===================================
'Check that call duration was entered

    iCallTime = 6
    
    If tmrMain.Enabled = True Then tmrMain.Enabled = False
    
    If Not txtMinutes.Text = "Enter Minutes Here" And Not txtMinutes.Text = "" Then
        If CLng(CDate(txtMinutes.Text) * 24 * 60) > 480 Then
            iCallTime = CInt(txtMinutes.Text)
        Else
            iCallTime = CInt(CDate(txtMinutes.Text) * 24 * 60)
        End If
    Else
        MsgBox "Please enter the call duration.", vbExclamation, "FMain::cmdUpdate"
        txtMinutes.Text = "Enter Minutes Here"
        Exit Sub
    End If

'===================================
'Check that a note was entered

    If txtCallNote(0).Text = "" And txtCallNote(1).Text = "" Then
        MsgBox "Enter call notes.", vbExclamation, "FMain::cmdUpdate"
        Exit Sub
    End If
    If Len(Trim(txtCallNote(0).Text)) > 0 Then
        sNote = sNote & "Questions:" & vbCrLf & txtCallNote(0).Text & vbCrLf
    End If
    If Len(Trim(txtCallNote(1).Text)) > 0 Then
        sNote = sNote & "Answers:" & vbCrLf & txtCallNote(1).Text
    End If
    
    If lblCaseID.Caption = "Enter ID" Then lblCaseID.Caption = "0"

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

If fEditCaseID Then
    If BLServer.AddCall2(iCustomerID, iContactID, iCallCodeID, iProductID, iEmplID, lblDate.Caption, sNote, iCallTime, CLng(lblCaseID.Caption), lCurrentRecordID) = 0 Then
        MsgBox "Error adding call", vbCritical, "FMain::cmdEditCall"
        Exit Sub
    End If
Else
    If BLServer.AddCall(iCustomerID, iContactID, iCallCodeID, iProductID, iEmplID, lblDate.Caption, sNote, iCallTime, CLng(lblCaseID.Caption)) = 0 Then
        MsgBox "Error adding call", vbCritical, "FMain::cmdUpdate"
        Exit Sub
    End If
End If

'Success so reset the form
txtCallHistory.Text = sText
GetCompanyHistory iCustomerID, txtCallHistory
ResetControls

Set BLServer = Nothing

Exit Sub

ErrorHandler:
    Set BLServer = Nothing
    MsgBox "Failed to add call to database : " & Err.Description, vbCritical, "CL-cmdUpdate::CLICK"
End Sub
'##ModelId=3A0F61DD03AC
Sub ResetControls()

fEditCaseID = False

With Me
    With .cmdEditCall
        .Enabled = False
        .Caption = "   Reset         Form"
    End With
    .cmdUpdate.Caption = "   Add   New Call"
    .lblDate = Format(Now(), "mmm-dd-yy")
    .txtMinutes.Text = "Enter Minutes Here"
    .txtEntry(clCOMPANY).Text = ""
    .txtEntry(clCONTACT).Text = ""
    .txtEntry(2).Text = ""
    .txtEntry(3).Text = ""
    .txtCallNote(0).Text = ""
    .txtCallNote(1).Text = ""
    .lblCaseID.Caption = "0"
    
    .lstItem(2).ListIndex = -1
    .lstItem(3).ListIndex = -1

End With

End Sub

Private Sub Form_Activate()
On Error GoTo FMainEHandler
    If Not cmdForceTVWLoad.Enabled Then
        cmdForceTVWLoad.Enabled = True
    End If
Exit Sub

FMainEHandler:
MsgBox Err.Number & ": " & Err.Description & vbCrLf & Err.Source, vbCritical, "Error FormActivate"
End Sub

Private Sub Form_Click()
'Me.Refresh
'MsgBox Me.WindowState
End Sub

Private Sub Form_GotFocus()
Screen.MousePointer = vbDefault
End Sub

'##ModelId=3A0F61DE0032
Private Sub Form_Initialize()
On Error GoTo FMainHandler

InitForm
InitLists

Exit Sub

FMainHandler:
MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & Err.Source, vbCritical, "Error Form_Initialize"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
MsgBox KeyAscii
End Sub

Private Sub Form_Load()
On Error GoTo FMainHandler

If LoadCurrentCustomers = 0 Then
    MsgBox "Failed loading current customers", vbCritical, "FMain2::Form_Load"
End If

Exit Sub

FMainHandler:
MsgBox Err.Number & ": " & Err.Description & vbCrLf & Err.Source, vbCritical, "Error Form_Load"
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
Dim iResult As Integer

xScreenRes = GetSystemMetrics(0)
yScreenRes = GetSystemMetrics(1)

iResult = ResizeControls
'    MsgBox "did something"
'Me.txtCallNote(0).Text = "h=" & Me.Height & vbCrLf & "sh=" & Me.ScaleHeight & vbCrLf & "w=" & Me.Width & vbCrLf & "sw=" & Me.ScaleWidth & vbCrLf & "sbb=" & Me.sbMain.Top + Me.sbMain.Height
'iResult = ResizeControls

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
'    txtEntry(Index).Text = lstItem(Index).Text
    Select Case Index
        Case 0 'Company
            'Clear the contact list box
'            lstItem(1).Clear
            'Clear the contact text box
            txtEntry(Index).Text = lstItem(Index).Text
            txtEntry(clCONTACT).Text = ""
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
            If Not cboEmployees.ListIndex = 0 Then
                cmdEditContact.Enabled = True
                cmdEditContact.Caption = "Edit Company"
                cmdNewContact.Enabled = True
            End If
            cmdEditContact.Caption = "Edit Contact"

        Dim lEmplID As Long
            With lstItem(Index)
                iSelectedContactID = .ItemData(lstItem(Index).ListIndex)
        '        MsgBox .ItemData(lstItem(Index).ListIndex)
        '        If Not .HitTest(X, Y) Is Nothing Then
        '            .SelectedItem = .HitTest(X, Y)
        '            If Not cboEmployees.ItemData(cboEmployees.ListIndex) = 0 Then
        '                If Not (.SelectedItem.Parent Is Nothing) Then
        '                    'Get contact's ID
        '                    iSelectedContactID = .ItemData
        '                    'CLng(Right(.SelectedItem.Key, Len(.SelectedItem.Key) - 1))
        '                    fDrag = True
        '                End If
        '            End If
        '        End If
            End With
            txtEntry(Index).Text = lstItem(Index).Text & ":::c" & iSelectedContactID
        Case 2 'Product
        Case 3 'Call code
    End Select
End If
End Sub

Private Sub mfgHistory_Click()
'MsgBox mfgHistory.CellBackColor
Dim iCounter As Integer
Dim iCounter2 As Integer
Dim iRow As Integer
Dim iTempRow As Integer
Dim lCaseID As Long
Dim sTemp As String
Dim sTempText As String

On Error GoTo ErrorHandler

With Me.mfgHistory
    iRow = .Row
    .Col = 0
    .Row = 1
    Me.txtCallNote(0).Text = ""
    Me.txtCallNote(1).Text = ""
    Me.txtMinutes.Text = ""
    Me.lblCaseID.Caption = "0"
'    Me.cmdEditCall.Caption = "Clear Call"
    
    If Not .Text = "No Data" Then
        .Row = iRow
        .Row = 0
        For iCounter = 0 To .Cols - 1
            .Col = iCounter
            sTemp = .Text
            .Row = iRow
            sTempText = .Text
            Select Case sTemp
                Case "CaseID"
'                    Me.lblCaseID.Caption = Left(sTempText, InStr(1, sTempText, " (") - 1)
                    Me.lblCaseID.Caption = sTempText
                Case "RecordID"
                    lCurrentRecordID = CLng(sTempText)
                Case "Duration"
                    Me.txtMinutes.Text = sTempText
                Case "Date"
                    Me.lblDate.Caption = Format(sTempText, "dd-mmm-yy")
                Case "Note"
            'Check for Answers in note
                    If InStr(1, sTempText, "Answers:") > 0 Then
                        Me.txtCallNote(1).Text = Mid(sTempText, (InStr(1, sTempText, "Answers:") + 10))
                        If InStr(1, sTempText, "Answers:") > 1 Then
                            sTempText = Left(sTempText, (InStr(1, sTempText, "Answers:") - 1))
                        Else
                            sTempText = ""
                        End If
                    End If
                    
            'Check for Questions in note
                    If InStr(1, sTempText, "Questions:") > 0 Then
                        Me.txtCallNote(0).Text = Mid(sTempText, (InStr(1, sTempText, "Questions:") + 12), Len(sTempText) - (InStr(1, sTempText, "Questions:") + 13))
                        If InStr(1, sTempText, "Questions:") > 1 Then
                            sTempText = Left(sTempText, (InStr(1, sTempText, "Questions:") - 3))
                        Else
                            sTempText = ""
                        End If
                    End If
                    
                    If (Me.txtCallNote(0).Text & Me.txtCallNote(1).Text) = "" Then
                        Me.txtCallNote(0).Text = sTempText
                    End If
            'Check for contact in note
                    If InStr(1, sTempText, "Contact:") > 0 Then
                        Me.txtEntry(clCONTACT).Text = Mid(sTempText, InStr(1, sTempText, "Contact:") + 9) ', InStr(1, sTempText, Chr(13)) - (InStr(1, sTempText, "Contact:") + 9))
                        If InStr(1, sTempText, "Contact:") > 1 Then
                            sTempText = Left(sTempText, (InStr(1, sTempText, "Contact:") - 3))
                        Else
                            sTempText = ""
                        End If
                    End If
                    
            'Check for company in note
                    If InStr(1, sTempText, "Company:") > 0 Then
                        Me.txtEntry(clCOMPANY).Text = Mid(sTempText, InStr(1, sTempText, "Company:") + 9)
                    End If
                Case "Product"
                    For iCounter2 = 0 To Me.lstItem(2).ListCount - 1
                        If Me.lstItem(2).List(iCounter2) = sTempText Then
                            Me.lstItem(2).ListIndex = iCounter2
                            Exit For
                        End If
'                        MsgBox Me.lstItem(2).List(iCounter2)
                    Next iCounter2
'                    Me.lstItem(0).Text
                Case "Call Type"
'                    Me.optCallCode(Mid(sTempText, InStr(1, sTempText, "(") + 1, Len(sTempText) - InStr(1, sTempText, "(") - 1)).Value = True
                Case "Call Code ID"
                    Me.optCallCode(sTempText).Value = True
                Case "Contact"
'                    If Right(sTempText, 4) = "(10)" Then
'                    Else
'                        For iCounter2 = 0 To Me.lstItem(1).ListCount - 1
'                            If Me.lstItem(1).ItemData(iCounter2) = CLng(Mid(sTempText, InStr(1, sTempText, " (") + 2, Len(sTempText) - (InStr(1, sTempText, "(") + 1))) Then
'                                Me.lstItem(1).ListIndex = iCounter2
'                            End If
'                        Next iCounter2
''                        Me.lstItem(1).ListIndex = Mid(sTempText, (InStr(1, sTempText, "(") + 1), InStr(1, sTempText, ")") - (InStr(1, sTempText, "(") + 1))
'                    End If
                Case "ContactID"
                    If sTempText = "10" Then
                    Else
                        For iCounter2 = 0 To Me.lstItem(1).ListCount - 1
                            If Me.lstItem(1).ItemData(iCounter2) = CLng(sTempText) Then
                                Me.lstItem(1).ListIndex = iCounter2
                                If Not Me.tvwCustomers.SelectedItem Is Nothing Then
                                    tvwCustomers_Click
                                    lstItem_Click (1)
                                End If
                            End If
                        Next iCounter2
'                        Me.lstItem(1).ListIndex = Mid(sTempText, (InStr(1, sTempText, "(") + 1), InStr(1, sTempText, ")") - (InStr(1, sTempText, "(") + 1))
                    End If
                Case "PSG Engr"
            End Select
            .Row = 0
        Next iCounter
    End If
End With

With Me
    With .cmdEditCall
        If Me.cboEmployees.ItemData(Me.cboEmployees.ListIndex) <> 0 Then
        .Enabled = True
        .Caption = "   Reset         Form"
        End If
    End With
    .cmdUpdate.Caption = "   Edit   Old Call"
End With

fEditCaseID = True

Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::tvwCustomers"

End Sub

Private Sub tmrBLStatus_Timer()

If BLServer Is Nothing Then 'And rsGeneric Is Nothing Then
    shBLStatus.FillColor = vbRed
    Me.sbMain.Panels(3).Text = ""
Else
    shBLStatus.FillColor = vbGreen
    Me.sbMain.Panels(3).Text = "BLServer connection ACTIVE"
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
'IF is required when clearing the selection
'Static sCurrentParent As String

On Error GoTo tvwCustomersErrorHandler

'fEditCaseID = False
'Me.cmdEditCall.Enabled = False

If Not cboEmployees.ListIndex = 0 Then
    cmdEditContact.Enabled = True
    cmdEditContact.Caption = "Edit Company"
    cmdNewContact.Enabled = True
End If

With tvwCustomers.SelectedItem
    If .Parent Is Nothing Then
        'This is a CUSTOMER (parent)
        If sCurrentParent <> .Key Then
            LoadOneCustomer clCONTACT3, Right(.Key, Len(.Key) - 1)
            GetCompanyHistory Right(.Key, Len(.Key) - 1), txtCallHistory ', Right(.Key, Len(.Key) - 1)
            If Not fEditCaseID Then lblCaseID.Caption = "0"
        End If
        .Expanded = True
        txtEntry(clCOMPANY).Text = .Text & " :::" & .Key
        txtEntry(clCONTACT).Text = ""
    Else
        'It is a CONTACT (child)
        txtEntry(clCOMPANY).Text = .Parent.Text
        txtEntry(clCONTACT).Text = .Text & " :::" & .Key
        If sCurrentParent <> .Parent.Key Then
            GetCompanyHistory Right(.Parent.Key, Len(.Parent.Key) - 1), txtCallHistory
            If Not fEditCaseID Then lblCaseID.Caption = "0"
        End If
    End If
    If InStr(1, .Key, "p") Then sCurrentParent = .Key Else sCurrentParent = .Parent.Key
'    If Left(.Key, 1) = "p" Then sCurrentParent = .Key Else sCurrentParent = .Parent.Key
End With
Exit Sub
tvwCustomersErrorHandler:
Select Case Err.Number
    Case 3021
        MsgBox tvwCustomers.SelectedItem.Text & " has no contacts, add a contact before proceeding.", vbInformation, "CL-FMain::tvwCustomers"
    Case Else
        MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::tvwCustomers"
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

Private Sub tvwCustomers_KeyUp(KeyCode As Integer, Shift As Integer)
Dim iCounter As Integer
'MsgBox KeyCode
Select Case KeyCode
    Case 40
        For iCounter = Me.tvwCustomers.SelectedItem.Index To Me.tvwCustomers.Nodes.Count '
            If InStr(1, UCase(Me.tvwCustomers.Nodes.Item(iCounter).Text), UCase(txtEntry(clCOMPANY).Text)) > 0 Then
                Me.tvwCustomers.Nodes.Item(iCounter).Selected = True
                Exit For
            End If
        Next iCounter
    Case 38
        For iCounter = Me.tvwCustomers.SelectedItem.Index To 1 Step -1
            If InStr(1, UCase(Me.tvwCustomers.Nodes.Item(iCounter).Text), UCase(txtEntry(clCOMPANY).Text)) > 0 Then
                Me.tvwCustomers.Nodes.Item(iCounter).Selected = True
                Exit For
            End If
        Next iCounter
End Select
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
On Error GoTo WrongValue

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

'Set controls if Case ID was entered
If lblCaseID.Caption <> "" And lblCaseID.Caption <> "Enter ID" Then
    If tvwCustomers.SelectedItem Is Nothing Then
        lblCaseID.Caption = txtEnterCaseID.Text
'        lblCaseID.Visible = True
'        txtEnterCaseID.Visible = False
'        txtEnterCaseID.Text = "0"
        Exit Sub
    End If
'        With tvwCustomers.SelectedItem
'            If Left(.Key, 1) = "p" Then
'                'This is a CUSTOMER (parent)
'                lCompanyKey = CLng(Right(.Key, Len(.Key) - 1))
'            Else
'                'It is a CONTACT (child)
'                lCompanyKey = CLng(Right(.Parent.Key, Len(.Parent.Key) - 1))
'            End If
'        End With
Else 'reset the ID
    lblCaseID.Caption = "0"
End If

lblCaseID.Visible = True
txtEnterCaseID.Visible = False
txtEnterCaseID.Text = "0"

'Reset the controls if the ID is not showing
If lblCaseID.Visible = False Then
    lblCaseID.Caption = txtEnterCaseID.Text
    lblCaseID.Visible = True
    txtEnterCaseID.Visible = False
    txtEnterCaseID.Text = "0"
    fEditCaseID = False
End If
    
'Not editing an existing case so find history for case ID
If Not fEditCaseID Then
    lTemp = FindCaseID(lblCaseID.Caption)
    sTemp = "p" & CStr(lTemp)
    
    If lTemp > 0 Then
        With tvwCustomers.Nodes.Item(sTemp)
            .Selected = True
            LoadOneCustomer clCOMPANYBYID, lTemp
            LoadOneCustomer clCONTACT3, Right(.Key, Len(.Key) - 1)
            .Expanded = True
            GetCompanyHistory lTemp, txtCallHistory, lblCaseID.Caption
        End With
    End If
End If

End Sub

Private Sub txtEntry_Change(Index As Integer)
If Index = 0 Then 'company name changed by user
'    If txtEntry(Index).Text <> tvwCustomers.SelectedItem.Text Then
'        Me.tvwCustomers.SelectedItem.Selected = False
'    End If
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

Private Sub txtEntry_GotFocus(Index As Integer)
    txtEntry(Index).SelStart = 0
    txtEntry(Index).SelLength = Len(txtEntry(Index).Text)
End Sub

'##ModelId=3A0F61DF005B
Private Sub txtEntry_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Call FindListItem(lstItem(Index), txtEntry(Index))
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
Dim NewNode As Node
Dim iCounter As Integer

If Index = 0 Then
    RefreshCustomers clREFRESHCOMPANIES, txtEntry(Index).Text
End If

For iCounter = 1 To Me.tvwCustomers.Nodes.Count
    If InStr(1, UCase(Me.tvwCustomers.Nodes.Item(iCounter).Text), UCase(txtEntry(Index).Text)) > 0 Then
        Me.tvwCustomers.Nodes.Item(iCounter).Selected = True
        Exit For
    End If
Next iCounter

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

'        For Items = 0 To 4
'            Set Entry = Entries.Item(Number - Items)
'            EntryText = Entry.EIndex & " : " & Entry.EmplID & " : " & Entry.EDate & " : " & Entry.ETime
'            If Entries.Count >= Items + 1 Then lblEntryData(Items).Caption = EntryText
'            If Entries.Count >= Items + 1 Then lblEntryData.Caption = EntryText
'            If Entries.Count - 1 <= Items Then Exit For
'        Next
        
        If Entries.Count > 5 Then MaxItems = 4 Else MaxItems = Entries.Count - 1
        
'        For Items = 0 To MaxItems
'            cmdDelCall(Items).Enabled = True
'            If Entries.Count - 1 <= Items Then Exit For
'        Next
    
    End If 'Is there a call to enter
'vscCalls.Max = Entries.Count
'vscCalls.Value = Entries.Count
'lblTotalCalls.Caption = "Total Calls Entered = " & CStr(Entries.Count)

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
'Private Sub vsccalls_Change()
'Dim MaxItems As Integer
'Dim Items As Integer
'Dim EntryText As String
'
'    If Entries.Count < 5 Then vscCalls.Min = Entries.Count
'    If Entries.Count >= 5 Then vscCalls.Min = 5
'    If Entries.Count = 0 Then Exit Sub
'    If Entries.Count >= 5 Then MaxItems = 4 Else MaxItems = Entries.Count - 1
'
'    For Items = 0 To MaxItems
'        Set Entry = Entries.Item(vscCalls.Value - Items)
'        EntryText = Entry.EIndex & " : " & Entry.EmplID & " : " & Entry.EDate & " : " & Entry.ETime
'        If Entries.Count >= Items + 1 Then lblEntryData(Items).Caption = EntryText
'        If Entries.Count - 1 <= Items Then Exit For
'    Next
'End Sub
Private Sub LoadLists(iIndex As Integer, ByRef oData As Object, Optional Filter As String)
'Routine fills a collection with data for a listbox/combobox control
'iIndex is the action constant
'oData is a collection
'
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

If Filter <> "" Then
    Set rsGeneric = BLServer.GetLbData2(iIndex, Filter)
Else
    Set rsGeneric = BLServer.GetLbData2(iIndex)
End If
If rsGeneric Is Nothing Then
    MsgBox "Recordset not created" & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadLists"
    Exit Sub
End If

rsGeneric.MoveLast
LoadProgressMax = rsGeneric.AbsolutePosition
LoadProgressActive = True

'frmSplash.pgbLoadTVW.Value = frmSplash.pgbLoadTVW.Min
'frmSplash.pgbLoadTVW.Max = rsGeneric.AbsolutePosition


rsGeneric.MoveFirst
Do While Not rsGeneric.EOF
    Select Case iIndex
        Case clLINK
            If Not (IsNull(rsGeneric!ID) Or IsNull(rsGeneric!CompanyID) Or IsNull(rsGeneric!ContactID)) Then
                If oData.Add(rsGeneric!ID, rsGeneric!CompanyID, rsGeneric!ContactID) Is Nothing Then MsgBox "Error adding Customer", , "CL-FMain::LoadLists"
            End If
        Case clCONTACT3
            If Not (IsNull(rsGeneric!ChildID) Or IsNull(rsGeneric!sName)) Then
                If oData.Add(rsGeneric!sName, Right(rsGeneric!ChildID, Len(rsGeneric!ChildID) - 1)) Is Nothing Then MsgBox "Error adding Customer", , "FMain-InitLists"
            End If
        Case Else
            If Not (IsNull(rsGeneric!ID) Or IsNull(rsGeneric!sName)) Then
                If oData.Add(rsGeneric!sName, rsGeneric!ID) Is Nothing Then MsgBox "Error adding Customer", , "FMain-InitLists"
            End If
    End Select
    LoadProgress = rsGeneric.AbsolutePosition
'    frmSplash.pgbLoadTVW.Value = rsGeneric.AbsolutePosition
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
        MsgBox "Unable to access Server" & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadCustomer"
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
    Case clCOMPANYBYDATE
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
        MsgBox "Unable to access Server" & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadCustomer2"
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
        Case clCOMPANYBYID
            Set NewNode = tvwCustomers.Nodes.Add(, , sChildID, sNodeName, CInt(rsGeneric!cType), CInt(rsGeneric!cType) + 20)
        Case clCONTACTBYID
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
Dim iNumberOfRecords As Integer

On Error GoTo LoadAllCustomersErrorHandler

frmSplash.MousePointer = vbHourglass
iCounter = 0

'Make sure a server exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    shBLStatus.FillColor = vbGreen
    If BLServer Is Nothing Then
        shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server" & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadAllCustomers"
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

rsGeneric.MoveLast
iNumberOfRecords = rsGeneric.AbsolutePosition
rsGeneric.MoveFirst

tvwCustomers.Sorted = False

If frmSplash.Visible = False Then
    tvwCustomers.Nodes.Clear
    Do While Not rsGeneric.EOF
        sParentID = (rsGeneric!ParentID)
        sChildID = (rsGeneric!ChildID)
        sNodeName = rsGeneric!sName
'        sNodeName = rsGeneric!Name
        iType = CInt(rsGeneric!cType)
        With tvwCustomers.Nodes
            Select Case iIndex
                Case clCOMPANYBYDATE
                    Set NewNode = .Add(, , sChildID, sNodeName, iType, iType + 20)
                    NewNode.Expanded = False
                    Set NewNode = Nothing
'                        sbMain.Panels(3).Text = "Loading Companies ..." & sParentID
'                        frmSplash.Label1.Caption = "Loading Companies ..." & sParentID
'                        frmSplash.txtJWVTest.Text = frmSplash.txtJWVTest.Text & sParentID & " :: " & sChildID & " :: " & sNodeName & " :: " & iType & vbCrLf
                Case clCONTACT
                    Set NewNode = .Add(sParentID, tvwChild, sChildID, sNodeName, iType, iType + 21)
                    NewNode.Expanded = False
                    Set NewNode = Nothing
'                        sbMain.Panels(3).Text = "Loading Contact ..." & sChildID & " to " & sParentID
'                        frmSplash.Label1.Caption = "Loading Contacts ..." & sChildID
            End Select
        End With
'            frmSplash.pgbLoadTVW.Value = rsGeneric.AbsolutePosition
'            frmSplash.Refresh
        sbMain.Panels(3).Text = Format(CStr(rsGeneric.AbsolutePosition / iNumberOfRecords * 100), "#0") & "% Loading Companies ..." & sParentID
        rsGeneric.MoveNext
        DoEvents
    Loop
    sbMain.Panels(3).Text = ""
Else 'on the splash screen
    Do While Not rsGeneric.EOF
        sParentID = (rsGeneric!ParentID)
        sChildID = (rsGeneric!ChildID)
'        sNodeName = rsGeneric!sName
        sNodeName = rsGeneric!Name
        iType = CInt(rsGeneric!cType)
        With tvwCustomers.Nodes
            Select Case iIndex
                Case clCOMPANYBYDATE
                    Set NewNode = .Add(, , sChildID, sNodeName, iType, iType + 20)
                    NewNode.Expanded = False
                    frmSplash.Label1.Caption = "Loading Companies ..." & sParentID
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

Function LoadCurrentCustomers() As Integer
Dim rsGeneric As ADODB.Recordset
Dim NewNode As Node
Dim sParentID As String
Dim sChildID As String
Dim sNodeName As String
Dim iCounter As Long
Dim iType As Integer
Dim iCompanyImage As Integer
Dim iNumberOfRecords As Integer

On Error GoTo LoadCurrCustomersErrorHandler

LoadCurrentCustomers = 0

frmSplash.MousePointer = vbHourglass
iCounter = 0

'Make sure a server exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    shBLStatus.FillColor = vbGreen
    If BLServer Is Nothing Then
        shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server" & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadCurrentCustomers"
        Exit Function
    End If
End If

'Get the customer data from the server
Set rsGeneric = BLServer.GetLbData2(clNEWCOMPANIES)
If rsGeneric Is Nothing Then
    MsgBox "Recordset not created"
    If Not BLServer Is Nothing Then Set BLServer = Nothing
    Exit Function
End If

'Fill the treeview
If frmSplash.Visible = True Then
    rsGeneric.MoveLast
    LoadProgressActive = True
    LoadProgressMax = rsGeneric.AbsolutePosition
End If

'Load customers
If Not (rsGeneric.EOF And rsGeneric.BOF) Then
    rsGeneric.MoveLast
    iNumberOfRecords = rsGeneric.AbsolutePosition
    rsGeneric.MoveFirst
    
    tvwCustomers.Sorted = False
    
'    If fQuery.Visible = False Then 'This is not for queries
'        If frmSplash.Visible = False Then 'This is not during startup
            tvwCustomers.Nodes.Clear
            Do While Not rsGeneric.EOF
                sParentID = (rsGeneric!ParentID)
                sChildID = (rsGeneric!ChildID)
                sNodeName = rsGeneric!sName
                iType = CInt(rsGeneric!cType)
                With tvwCustomers.Nodes
                    Set NewNode = .Add(, , sChildID, sNodeName, iType, iType + 20)
                    NewNode.Expanded = False
                    Set NewNode = Nothing
                End With
                sbMain.Panels(3).Text = Format(CStr(rsGeneric.AbsolutePosition / iNumberOfRecords * 100), "#0") & "% Loading Companies ..." & sParentID
'MsgBox "debug", vbInformation, "Temp Window"
                rsGeneric.MoveNext
                DoEvents
            Loop
            sbMain.Panels(3).Text = ""
'        Else 'on the splash screen
'            Do While Not rsGeneric.EOF
'                sParentID = (rsGeneric!ParentID)
'                sChildID = (rsGeneric!ChildID)
'                sNodeName = rsGeneric!sName
'                iType = CInt(rsGeneric!cType)
'                With tvwCustomers.Nodes
'                    Set NewNode = .Add(, , sChildID, sNodeName, iType, iType + 20)
'                    NewNode.Expanded = False
'                    Set NewNode = Nothing
'                    frmSplash.Label1.Caption = "Loading Companies ..." & sParentID
'                    frmSplash.txtJWVTest.Text = frmSplash.txtJWVTest.Text & sParentID & " :: " & sChildID & " :: " & sNodeName & " :: " & iType & vbCrLf
'                End With
'                LoadProgress = rsGeneric.AbsolutePosition
'                rsGeneric.MoveNext
'                frmSplash.Refresh
'                DoEvents
'            Loop
'        End If
'    Else
'        fQuery.txtCallHistory.Text = ""
'        iCounter = 0
'        Do While Not rsGeneric.EOF
'            fQuery.txtCallHistory.Text = fQuery.txtCallHistory.Text & vbCrLf & _
'                rsGeneric!sName & " (" & rsGeneric!ParentID & ")"
'            rsGeneric.MoveNext
'            DoEvents
'            iCounter = iCounter + 1
'        Loop
'        fQuery.txtCallHistory.Text = fQuery.txtCallHistory.Text & vbCrLf & vbCrLf & _
'        "Number of entries = " & iCounter
'    End If
Else
    Me.sbMain.Panels(3).Text = "No current records"
End If
iCompanyImage = -1

''Get the customer data from the server
'Set rsGeneric = BLServer.GetLbData2(clNEWCOMPANIESNOCALLS)
'If rsGeneric Is Nothing Then
'    MsgBox "Recordset not created"
'    If Not BLServer Is Nothing Then Set BLServer = Nothing
'    Exit Sub
'End If
'
''Load customers
'If Not (rsGeneric.EOF And rsGeneric.BOF) Then
'    rsGeneric.MoveLast
'    iNumberOfRecords = rsGeneric.AbsolutePosition
'    rsGeneric.MoveFirst
'
'    tvwCustomers.Sorted = False
'
'    tvwCustomers.Nodes.Clear
'    Do While Not rsGeneric.EOF
'        sParentID = (rsGeneric!ParentID)
'        sChildID = (rsGeneric!ChildID)
'        sNodeName = rsGeneric!sName
'        iType = CInt(rsGeneric!cType)
'        With tvwCustomers.Nodes
'            Set NewNode = .Add(, , sChildID, sNodeName, iType, iType + 20)
'            NewNode.Expanded = False
'            Set NewNode = Nothing
'        End With
'        sbMain.Panels(3).Text = Format(CStr(rsGeneric.AbsolutePosition / iNumberOfRecords * 100), "#0") & "% Loading Companies ..." & sParentID
'        rsGeneric.MoveNext
'        DoEvents
'    Loop
'    sbMain.Panels(3).Text = ""
'Else
'    Me.sbMain.Panels(3).Text = "No current records"
'End If
'iCompanyImage = -1

'Set NewNode = Nothing
Set rsGeneric = Nothing
Set BLServer = Nothing

tvwCustomers.Sorted = True
LoadCurrentCustomers = 1

Exit Function

LoadCurrCustomersErrorHandler:

    MsgBox "Actual Node Count: " & SendMessage(tvwCustomers.hwnd, _
           TVM_GETCOUNT, 0, ByVal 0)
    MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadAllCustomers"
Set rsGeneric = Nothing
'Resume Next
frmSplash.MousePointer = vbDefault
End Function

Sub LoadOneCustomer(iIndex As Integer, Optional Filter As Long)
'Routine adds/updates a company and/or contact
'
'iIndex = update action constant
'Filter = company or contact PKId
'
'Company changes are to the treeview
'Contact changes are to the lstItem(1) listbox

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
Me.lstItem(1).Clear

Screen.MousePointer = vbHourglass

'Make sure a server exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    shBLStatus.FillColor = vbGreen
    If BLServer Is Nothing Then
        shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server, notify administrator.", vbCritical, "CL-FMain::LoadOneCustomer"
        Exit Sub
    End If
End If

'Get the customer data from the server
Set rsGeneric = BLServer.GetLbData2(iIndex, CStr(Filter))
If rsGeneric Is Nothing Or rsGeneric.RecordCount = 0 Then
    If iIndex = clCONTACT3 Then
    Else
        MsgBox "Unable to get information, notify administrator.", vbCritical, "CL-Main::LoadOneCustomer"
    End If
    Screen.MousePointer = vbDefault
    Set rsGeneric = Nothing
    Set BLServer = Nothing
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
        Case clCOMPANYBYDATE
            Set NewNode = .Add(, , sChildID, sNodeName, iType, iType + 20)
            NewNode.Expanded = False
            sbMain.Panels(3).Text = "Loading Companies ..." & sParentID
        'Edit an existing company node
        Case clCOMPANYBYID
            Set NewNode = .Add(, , sChildID, sNodeName, iType, iType + 20)
            NewNode.Expanded = False
            sbMain.Panels(3).Text = "Loading Companies ..." & sParentID
            With .Item(sParentID)
                .Text = sNodeName
                .Image = iType
                .SelectedImage = iType + 20
            End With
        'Add a new contact node
        Case clCONTACT
'            Set NewNode = .Add(sParentID, tvwChild, sChildID, sNodeName, iType, iType + 21)
'            NewNode.Expanded = False
            sbMain.Panels(3).Text = "Loading Contact ..." & sChildID & " to " & sParentID
            Me.lstItem(1).AddItem sNodeName ', Right(sChildID, Len(sChildID) - 1)
            Me.lstItem(1).ItemData(lstItem(1).NewIndex) = CLng(Right(rsGeneric!ChildID, Len(rsGeneric!ChildID) - 1))
        'Edit an existing contact node
'        Case clCONTACTBYID
'            With .Item(sChildID)
'                .Text = sNodeName
'                .Image = iType
'                .SelectedImage = CInt(rsGeneric!cType) + 21
'                tvwCustomers.Refresh
'            End With
        'Refresh the list of contacts (delete them then reload)
        Case clCONTACT3, clCONTACTBYID
            lstItem(1).Clear
            Do While Not rsGeneric.EOF
                lstItem(1).AddItem rsGeneric!sName
                lstItem(1).ItemData(lstItem(1).NewIndex) = CLng(Right(rsGeneric!ChildID, Len(rsGeneric!ChildID) - 1))
                rsGeneric.MoveNext
            Loop
        'Add a new company node
        Case clCOMPANYUPDATE
            With .Item(sParentID)
                .Text = sNodeName
                .Image = iType
                .SelectedImage = iType + 20
            End With
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

Sub UpdateOneCustomer(iIndex As Integer, Optional Filter As Long)
'Routine adds/updates a company and/or contact
'
'iIndex = update action constant
'Filter = company or contact PKId
'
'Company changes are to the treeview
'Contact changes are to the lstItem(1) listbox

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
Me.lstItem(1).Clear

Screen.MousePointer = vbHourglass

'Make sure a server exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    shBLStatus.FillColor = vbGreen
    If BLServer Is Nothing Then
        shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server, notify administrator.", vbCritical, "CL-FMain::LoadOneCustomer"
        Exit Sub
    End If
End If

'Get the customer data from the server
Set rsGeneric = BLServer.GetLbData2(iIndex, CStr(Filter))
If rsGeneric Is Nothing Or rsGeneric.RecordCount = 0 Then
    If iIndex = clCONTACT3 Then
    Else
        MsgBox "Unable to get information, notify administrator.", vbCritical, "CL-Main::LoadOneCustomer"
    End If
    Screen.MousePointer = vbDefault
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
        
'        Case clCOMPANYBYDATE 'Add a new company node
'            Set NewNode = .Add(, , sChildID, sNodeName, iType, iType + 20)
'            NewNode.Expanded = False
'            sbMain.Panels(3).Text = "Loading Companies ..." & sParentID
'
'        Case clCOMPANYBYID 'Edit an existing company node
'            Set NewNode = .Add(, , sChildID, sNodeName, iType, iType + 20)
'            NewNode.Expanded = False
'            sbMain.Panels(3).Text = "Loading Companies ..." & sParentID
'            With .Item(sParentID)
'                .Text = sNodeName
'                .Image = iType
'                .SelectedImage = iType + 20
'            End With
'
'        Case clCONTACT 'Add a new contact node
''            Set NewNode = .Add(sParentID, tvwChild, sChildID, sNodeName, iType, iType + 21)
''            NewNode.Expanded = False
'            sbMain.Panels(3).Text = "Loading Contact ..." & sChildID & " to " & sParentID
'            Me.lstItem(1).AddItem sNodeName ', Right(sChildID, Len(sChildID) - 1)
'            Me.lstItem(1).ItemData(lstItem(1).NewIndex) = CLng(Right(rsGeneric!ChildID, Len(rsGeneric!ChildID) - 1))
'
'        Case clCONTACTBYID 'Edit an existing contact node
'            With .Item(sChildID)
'                .Text = sNodeName
'                .Image = iType
'                .SelectedImage = CInt(rsGeneric!cType) + 21
'                tvwCustomers.Refresh
'            End With
        
        Case clCONTACT3, clCONTACTBYID 'Refresh the list of contacts (delete them then reload)
            lstItem(1).Clear
            Do While Not rsGeneric.EOF
                lstItem(1).AddItem rsGeneric!sName
                lstItem(1).ItemData(lstItem(1).NewIndex) = CLng(Right(rsGeneric!ChildID, Len(rsGeneric!ChildID) - 1))
                rsGeneric.MoveNext
            Loop
        
        Case clCOMPANYBYID, clCOMPANYUPDATE 'Edit company node
            With .Item(sParentID)
                .Text = sNodeName
                .Image = iType
                .SelectedImage = iType + 20
            End With
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
'txtCallHistory.ZOrder = 0
txtCallHistory.Visible = False

'Make sure a server exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
    shBLStatus.FillColor = vbGreen
    If BLServer Is Nothing Then
        shBLStatus.FillColor = vbRed
        MsgBox "Unable to access Server" & vbCrLf & "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::FindCaseID"
        Exit Function
    End If
End If

'Get the customer data from the server
Set rsGeneric = BLServer.GetLbData2(clCASEID, CLng(sCaseID))
If ((rsGeneric Is Nothing) Or (rsGeneric.RecordCount = 0)) Then
    MsgBox "Recordset not created"
    Exit Function
Else
    If (rsGeneric.BOF And rsGeneric.EOF) Then
'        txtCallHistory.ZOrder = 0
        txtCallHistory.Visible = True
        
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

Sub RefreshCustomers(iSearchIndex As Integer, Optional sSearchString As String)
Dim NewNode As Node
Dim sParentID As String
Dim sChildID As String
Dim sNodeName As String
Dim iCompanyImage As Integer
Dim iCounter As Long
Dim iType As Integer
Dim fNoNode As Boolean

On Error GoTo ErrorHandler
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
        MsgBox "Unable to access Server, notify administrator.", vbCritical, "CL-FMain::RefreshCustomers"
        Exit Sub
    End If
End If

'Get the customer data from the server
Set rsGeneric = BLServer.GetLbData2(iSearchIndex, sSearchString)
If rsGeneric Is Nothing Then
    MsgBox "Unable to get Contacts, notify administrator.", vbCritical, "CL-Main::RefreshCustomers"
    Exit Sub
End If
Set rsGeneric.ActiveConnection = Nothing
Set BLServer = Nothing

'Fill the treeview
If Not (rsGeneric.EOF And rsGeneric.BOF) Then
    rsGeneric.MoveFirst
    If CStr(rsGeneric!cType) = "" Then Exit Sub
    
    With tvwCustomers
        Select Case iSearchIndex
            'Add a new company node
            Case clREFRESHCOMPANIES
                    Do While Not rsGeneric.EOF
'                    If .Nodes.Count = 0 Then
'                        fNoNode = True
'                    End If
'                    For Each NewNode In .Nodes
'                        fNoNode = False
'                        If NewNode.Key = rsGeneric!ParentID Then Exit For
'                        fNoNode = True
'                    Next
'                    If fNoNode Then
                        Set NewNode = .Nodes.Add(, , rsGeneric!ChildID, rsGeneric!sName, CInt(rsGeneric!cType), CInt(rsGeneric!cType) + 20)
                        NewNode.Expanded = False
'                        Set NewNode = Nothing
'                    End If
                    rsGeneric.MoveNext
                Loop
            'Refresh the list of contacts (delete them then reload)
            Case clREFRESHCONTACTS
                    lstItem(1).Clear
                Do While Not rsGeneric.EOF
                    lstItem(1).AddItem rsGeneric!sName
                    lstItem(1).ItemData(lstItem(1).NewIndex) = CLng(Right(rsGeneric!ChildID, Len(rsGeneric!ChildID) - 1))
                    rsGeneric.MoveNext
                Loop
            'Add a new company node
            Case clCOMPANYUPDATE
        End Select
    End With
End If
DoEvents

iCompanyImage = -1

Set NewNode = Nothing
Set rsGeneric = Nothing
'Set BLServer = Nothing
Screen.MousePointer = vbDefault

Exit Sub

ErrorHandler:
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
    Case 35601, 35602, 94
    Case Else
        MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadOneCustomer"
End Select
'MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "CL-FMain::LoadOneCustomer"
Set rsGeneric = Nothing
Screen.MousePointer = vbDefault

End Sub
Sub DoSort()
    
    mfgHistory.Col = 0
    mfgHistory.ColSel = mfgHistory.Cols - 1
    mfgHistory.Sort = 1 ' Generic Ascending
    
End Sub
Private Sub mfghistory_DragDrop(Source As VB.Control, X As Single, Y As Single)
    If mfgHistory.Tag = "" Then Exit Sub
    mfgHistory.Redraw = False
    mfgHistory.ColPosition(Val(mfgHistory.Tag)) = mfgHistory.MouseCol
    DoSort
    mfgHistory.Redraw = True
End Sub

Private Sub mfghistory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mfgHistory.Tag = ""
    If mfgHistory.MouseRow <> 0 Then Exit Sub
    mfgHistory.Tag = str(mfgHistory.MouseCol)
    mfgHistory.Drag 1
End Sub


