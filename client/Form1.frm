VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1920
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   2778
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private UsingMouse As Boolean  ' Flag for using the Mouse in the Grid.

Private Sub Form_Load()  ' Set Control Property Values
   Text1.TabIndex = 1
   Text2.TabIndex = 2
   Text2.BorderStyle = 0
   Text3.TabIndex = 3
   MSFlexGrid1.Cols = 5
   MSFlexGrid1.Rows = 5
   MSFlexGrid1.TabStop = False
End Sub

Private Sub Text1_LostFocus()
   MSFlexGrid1.Col = 1
   MSFlexGrid1.Row = 1
End Sub

Private Sub Text2_GotFocus()
   MSFlexGrid1.Text = Text2.Text
   If MSFlexGrid1.Col >= MSFlexGrid1.Cols Then MSFlexGrid1.Col = 1
   ChangeCellText
End Sub

Private Sub MSFlexGrid1_EnterCell()  ' Assign cell value to the textbox
   Text2.Text = MSFlexGrid1.Text
End Sub

Private Sub MSFlexGrid1_LeaveCell()
' Assign textbox value to grid

   MSFlexGrid1.Text = Text2.Text
   Text2.Text = ""
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, _
                                  x As Single, y As Single)
   UsingMouse = True
   MSFlexGrid1.Text = Text2.Text
   ChangeCellText
End Sub

Private Sub Text2_LostFocus()
   If UsingMouse = True Then
      UsingMouse = False
      Exit Sub
   End If
   
   If MSFlexGrid1.Col <= MSFlexGrid1.Cols - 2 Then
      MSFlexGrid1.Col = MSFlexGrid1.Col + 1
      ChangeCellText
   Else
      If MSFlexGrid1.Row + 1 < MSFlexGrid1.Rows Then
        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
        MSFlexGrid1.Col = 1
        ChangeCellText
      End If
   End If
End Sub

Public Sub ChangeCellText() ' Move Textbox to active cell.
   Text2.Move MSFlexGrid1.Left + MSFlexGrid1.CellLeft, _
   MSFlexGrid1.Top + MSFlexGrid1.CellTop, _
   MSFlexGrid1.CellWidth, MSFlexGrid1.CellHeight
   Text2.SetFocus
   Text2.ZOrder 0
End Sub
