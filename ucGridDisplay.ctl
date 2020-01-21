VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl ucGridDisplay 
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10320
   PropertyPages   =   "ucGridDisplay.ctx":0000
   ScaleHeight     =   4230
   ScaleWidth      =   10320
   ToolboxBitmap   =   "ucGridDisplay.ctx":0068
   Begin MSFlexGridLib.MSFlexGrid fgDataGrid 
      Height          =   3015
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5318
      _Version        =   393216
   End
   Begin VB.CommandButton cmdControlButton 
      Caption         =   "Command1"
      Height          =   460
      Index           =   5
      Left            =   8520
      TabIndex        =   6
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdControlButton 
      Caption         =   "Command1"
      Height          =   460
      Index           =   4
      Left            =   6840
      TabIndex        =   5
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdControlButton 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      Height          =   460
      Index           =   3
      Left            =   5160
      TabIndex        =   4
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdControlButton 
      Caption         =   "Command1"
      Height          =   460
      Index           =   2
      Left            =   3480
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdControlButton 
      Caption         =   "Command1"
      Height          =   460
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdControlButton 
      Caption         =   "Command1"
      Height          =   460
      Index           =   0
      Left            =   25
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Image imgFunctionImage 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   9840
      Picture         =   "ucGridDisplay.ctx":037A
      Stretch         =   -1  'True
      Top             =   37
      Width           =   300
   End
   Begin VB.Label lblFunctionName 
      BackColor       =   &H80000010&
      Caption         =   " Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   10170
   End
End
Attribute VB_Name = "ucGridDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_BorderStyle = 0
Const m_def_DataMember = ""
Const m_def_WhatsThisHelpID = 0
Const m_def_ButtonHeight = 0
Const m_def_ButtonWidth = 0
Const m_def_ButtonLeft = 0
Const m_def_ButtonTop = 0
Const m_def_HeadingCaption = "0"
'Property Variables:
Dim m_ForeColor As Long
Dim m_Font As Font
Dim m_BorderStyle As Integer
'Dim m_DataFormat As IStdDataFormatDisp
Dim m_DataMember As String
Dim m_DataMembers As DataMembers
'Dim m_DataSource As DataSource
Dim m_ToolTipText As Control
Dim m_WhatsThisHelpID As Long
Dim m_ButtonHeight As Long
Dim m_ButtonWidth As Long
Dim m_ButtonLeft As Long
Dim m_ButtonTop As Long
Dim m_HeadingCaption As String
Dim m_HeadingFont As Font
Dim m_QueryRecordset As ADODB.Recordset

'Event Declarations:
Event Click() 'MappingInfo=cmdControlButton(0),cmdControlButton,0,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Change() 'MappingInfo=lblFunctionName,lblFunctionName,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event Hide() 'MappingInfo=UserControl,UserControl,-1,Hide
Attribute Hide.VB_Description = "Occurs when the control's Visible property changes to False."
Event InitProperties() 'MappingInfo=UserControl,UserControl,-1,InitProperties
Attribute InitProperties.VB_Description = "Occurs the first time a user control or user document is created."
Event ReadProperties(PropBag As PropertyBag) 'MappingInfo=UserControl,UserControl,-1,ReadProperties
Attribute ReadProperties.VB_Description = "Occurs when a user control or user document is asked to read its data from a file."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event Show() 'MappingInfo=UserControl,UserControl,-1,Show
Attribute Show.VB_Description = "Occurs when the control's Visible property changes to True."
Event WriteProperties(PropBag As PropertyBag) 'MappingInfo=UserControl,UserControl,-1,WriteProperties
Attribute WriteProperties.VB_Description = "Occurs when a user control or user document is asked to write its data to a file."
Event ButtonClick(Index As Integer) 'MappingInfo=cmdControlButton(0),cmdControlButton,0,Click

'=============================================================
'UserControl functions and properties (Start)
'
'
Private Sub UserControl_Initialize()
Dim iCount As Integer

With fgDataGrid
    .MergeCells = flexMergeRestrictColumns

    For iCount = 0 To .Cols - 1
        .MergeCol(iCount) = True     ' Allow merge on Columns 0 thru 3
        .ColAlignment(iCount) = 1
    Next iCount
End With

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    RaiseEvent ReadProperties(PropBag)
    
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.CurrentX = PropBag.ReadProperty("CurrentX", 0)
    UserControl.CurrentY = PropBag.ReadProperty("CurrentY", 0)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 4230)
    UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
    UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 10320)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
'    Set m_DataFormat = PropBag.ReadProperty("DataFormat", Nothing)
    Set m_DataMembers = PropBag.ReadProperty("DataMembers", Nothing)
'    Set m_DataSource = PropBag.ReadProperty("DataSource", Nothing)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Set m_ToolTipText = PropBag.ReadProperty("ToolTipText", Nothing)
    Set cmdControlButton(0).Font = PropBag.ReadProperty("ButtonFont", Ambient.Font)
    Set m_HeadingFont = PropBag.ReadProperty("HeadingFont", Ambient.Font)
    lblFunctionName.Alignment = PropBag.ReadProperty("Alignment", 0)
    lblFunctionName.AutoSize = PropBag.ReadProperty("AutoSize", False)
    lblFunctionName.Caption = PropBag.ReadProperty("Caption", " Label1")
    lblFunctionName.WordWrap = PropBag.ReadProperty("WordWrap", False)
    cmdControlButton(0).Caption = PropBag.ReadProperty("ButtonCaption", "Command1")
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_DataMember = PropBag.ReadProperty("DataMember", m_def_DataMember)
    m_WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", m_def_WhatsThisHelpID)
    m_ButtonHeight = PropBag.ReadProperty("ButtonHeight", m_def_ButtonHeight)
    m_ButtonWidth = PropBag.ReadProperty("ButtonWidth", m_def_ButtonWidth)
    m_ButtonLeft = PropBag.ReadProperty("ButtonLeft", m_def_ButtonLeft)
    m_ButtonTop = PropBag.ReadProperty("ButtonTop", m_def_ButtonTop)
    m_HeadingCaption = PropBag.ReadProperty("HeadingCaption", m_def_HeadingCaption)
End Sub

Private Sub UserControl_InitProperties()
    RaiseEvent InitProperties
    m_ForeColor = m_def_ForeColor
    Set m_Font = Ambient.Font
    m_BorderStyle = m_def_BorderStyle
    m_DataMember = m_def_DataMember
    m_WhatsThisHelpID = m_def_WhatsThisHelpID
    m_ButtonHeight = UserControl.Height
    m_ButtonWidth = UserControl.Width
'    m_ButtonLeft = UserControl.lblFunctionName
'    m_ButtonTop = UserControl.Top
    m_HeadingCaption = m_def_HeadingCaption
    Set m_HeadingFont = Ambient.Font
End Sub

'The Underscore following "Scale" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Scale
Public Sub Scale_(Optional X1 As Variant, Optional Y1 As Variant, Optional X2 As Variant, Optional Y2 As Variant)
    UserControl.Scale (X1, Y1)-(X2, Y2)
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
Dim iFormWidth As Integer
Dim iFormHeight As Integer

RaiseEvent Resize

If UserControl.Width < 5000 Then UserControl.Width = 5000
If UserControl.Height < 2000 Then UserControl.Height = 2000

iFormWidth = UserControl.Width - 30
iFormHeight = UserControl.Height - lblFunctionName.Height - cmdControlButton(0).Height

lblFunctionName.Width = iFormWidth
fgDataGrid.Width = iFormWidth

cmdControlButton(0).Width = (iFormWidth - 15 * 5) / 6
cmdControlButton(1).Width = cmdControlButton(0).Width
cmdControlButton(2).Width = cmdControlButton(0).Width
cmdControlButton(3).Width = cmdControlButton(0).Width
cmdControlButton(4).Width = cmdControlButton(0).Width
cmdControlButton(5).Width = cmdControlButton(0).Width

lblFunctionName.Top = 0
fgDataGrid.Top = lblFunctionName.Height
fgDataGrid.Height = iFormHeight

cmdControlButton(0).Top = fgDataGrid.Top + fgDataGrid.Height
cmdControlButton(1).Top = cmdControlButton(0).Top
cmdControlButton(2).Top = cmdControlButton(0).Top
cmdControlButton(3).Top = cmdControlButton(0).Top
cmdControlButton(4).Top = cmdControlButton(0).Top
cmdControlButton(5).Top = cmdControlButton(0).Top

cmdControlButton(0).Top = fgDataGrid.Top + fgDataGrid.Height
cmdControlButton(1).Left = cmdControlButton(0).Left + cmdControlButton(0).Width
cmdControlButton(2).Left = cmdControlButton(1).Left + cmdControlButton(1).Width + 15
cmdControlButton(3).Left = cmdControlButton(2).Left + cmdControlButton(2).Width + 15
cmdControlButton(4).Left = cmdControlButton(3).Left + cmdControlButton(3).Width + 15
cmdControlButton(5).Left = cmdControlButton(4).Left + cmdControlButton(4).Width + 15

imgFunctionImage.Height = lblFunctionName.Height - 125
imgFunctionImage.Width = imgFunctionImage.Height
imgFunctionImage.Top = lblFunctionName.Top + (lblFunctionName.Height - imgFunctionImage.Height) / 2
imgFunctionImage.Left = lblFunctionName.Left + lblFunctionName.Width - imgFunctionImage.Width - (lblFunctionName.Height - imgFunctionImage.Height) / 2

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    RaiseEvent WriteProperties(PropBag)
    
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("CurrentX", UserControl.CurrentX, 0)
    Call PropBag.WriteProperty("CurrentY", UserControl.CurrentY, 0)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 4230)
    Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
    Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
    Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 10320)
    
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
'    Call PropBag.WriteProperty("DataFormat", m_DataFormat, Nothing)
    Call PropBag.WriteProperty("DataMember", m_DataMember, m_def_DataMember)
    Call PropBag.WriteProperty("DataMembers", m_DataMembers, Nothing)
'    Call PropBag.WriteProperty("DataSource", m_DataSource, Nothing)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("HeadingCaption", m_HeadingCaption, m_def_HeadingCaption)
    Call PropBag.WriteProperty("HeadingFont", m_HeadingFont, Ambient.Font)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, Nothing)
    Call PropBag.WriteProperty("WhatsThisHelpID", m_WhatsThisHelpID, m_def_WhatsThisHelpID)
    
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    
    Call PropBag.WriteProperty("Alignment", lblFunctionName.Alignment, 0)
    Call PropBag.WriteProperty("AutoSize", lblFunctionName.AutoSize, False)
    Call PropBag.WriteProperty("Caption", lblFunctionName.Caption, " Label1")
    Call PropBag.WriteProperty("WordWrap", lblFunctionName.WordWrap, False)
    
    Call PropBag.WriteProperty("ButtonFont", cmdControlButton(0).Font, Ambient.Font)
    Call PropBag.WriteProperty("ButtonCaption0", cmdControlButton(0).Caption, "Command1")
    Call PropBag.WriteProperty("ButtonCaption1", cmdControlButton(1).Caption, "Command1")
    Call PropBag.WriteProperty("ButtonCaption2", cmdControlButton(2).Caption, "Command1")
    Call PropBag.WriteProperty("ButtonCaption3", cmdControlButton(3).Caption, "Command1")
    Call PropBag.WriteProperty("ButtonCaption4", cmdControlButton(4).Caption, "Command1")
    Call PropBag.WriteProperty("ButtonCaption5", cmdControlButton(5).Caption, "Command1")
    Call PropBag.WriteProperty("ButtonHeight", cmdControlButton(0).Height, m_def_ButtonHeight)
    Call PropBag.WriteProperty("ButtonWidth", cmdControlButton(0).Width, m_def_ButtonWidth)
    Call PropBag.WriteProperty("ButtonLeft", cmdControlButton(0).Left, m_def_ButtonLeft)
    Call PropBag.WriteProperty("ButtonTop", cmdControlButton(0).Top, m_def_ButtonTop)
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Hide()
    RaiseEvent Hide
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleX
Public Function ScaleX(ByVal Width As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
    ScaleX = UserControl.ScaleX(Width, FromScale, ToScale)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleY
Public Function ScaleY(ByVal Height As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
    ScaleY = UserControl.ScaleY(Height, FromScale, ToScale)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,CurrentY
Public Property Get CurrentY() As Single
    CurrentY = UserControl.CurrentY
End Property

Public Property Let CurrentY(ByVal New_CurrentY As Single)
    UserControl.CurrentY() = New_CurrentY
    PropertyChanged "CurrentY"
End Property

Public Property Let CurrentX(ByVal New_CurrentX As Single)
    UserControl.CurrentX() = New_CurrentX
    PropertyChanged "CurrentX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ContainerHwnd
Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,CurrentX
Public Property Get CurrentX() As Single
    CurrentX = UserControl.CurrentX
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'MappingInfo=UserControl,UserControl,-1,HasDC
Public Property Get HasDC() As Boolean
    HasDC = UserControl.HasDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Image
Public Property Get Image() As Picture
    Set Image = UserControl.Image
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleLeft
Public Property Get ScaleLeft() As Single
    ScaleLeft = UserControl.ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
    UserControl.ScaleLeft() = New_ScaleLeft
    PropertyChanged "ScaleLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As Integer
    ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
    UserControl.ScaleMode() = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleTop
Public Property Get ScaleTop() As Single
    ScaleTop = UserControl.ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
    UserControl.ScaleTop() = New_ScaleTop
    PropertyChanged "ScaleTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

'==================================================================
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
Private Sub UserControl_Show()
    RaiseEvent Show
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function DataMemberChanged(ByVal DataMember As String) As Boolean

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblFunctionName,lblFunctionName,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = "HeadingLabel"
    Alignment = lblFunctionName.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    lblFunctionName.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblFunctionName,lblFunctionName,-1,AutoSize
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
Attribute AutoSize.VB_ProcData.VB_Invoke_Property = "HeadingLabel"
    AutoSize = lblFunctionName.AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    lblFunctionName.AutoSize() = New_AutoSize
    PropertyChanged "AutoSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblFunctionName,lblFunctionName,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_ProcData.VB_Invoke_Property = "HeadingLabel"
    Caption = lblFunctionName.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblFunctionName.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=21,0,0,0
'Public Property Get DataFormat() As IStdDataFormatDisp
'    Set DataFormat = m_DataFormat
'End Property
'
'Public Property Set DataFormat(ByVal New_DataFormat As IStdDataFormatDisp)
'    Set m_DataFormat = New_DataFormat
'    PropertyChanged "DataFormat"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get DataMember() As String
Attribute DataMember.VB_Description = "Returns/sets a value that describes the DataMember for a data connection."
Attribute DataMember.VB_ProcData.VB_Invoke_Property = "DataSource"
    DataMember = m_DataMember
End Property

Public Property Let DataMember(ByVal New_DataMember As String)
    m_DataMember = New_DataMember
    PropertyChanged "DataMember"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=16,0,0,0
Public Property Get DataMembers() As DataMembers
Attribute DataMembers.VB_Description = "Returns a collection of data members to show at design time for this data source."
    Set DataMembers = m_DataMembers
End Property

Public Property Set DataMembers(ByVal New_DataMembers As DataMembers)
    Set m_DataMembers = New_DataMembers
    PropertyChanged "DataMembers"
End Property

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=22,0,0,0
'Public Property Get DataSource() As DataSource
'    Set DataSource = m_DataSource
'End Property
'
'Public Property Set DataSource(ByVal New_DataSource As DataSource)
'    Set m_DataSource = New_DataSource
'    PropertyChanged "DataSource"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=15,0,0,0
Public Property Get ToolTipText() As Object
Attribute ToolTipText.VB_Description = "Returns the control that has focus."
    Set ToolTipText = m_ToolTipText
End Property

'Public Property Set ToolTipText(ByVal New_ToolTipText As Control)
'    Set m_ToolTipText = New_ToolTipText
'    PropertyChanged "ToolTipText"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
    WhatsThisHelpID = m_WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
    m_WhatsThisHelpID = New_WhatsThisHelpID
    PropertyChanged "WhatsThisHelpID"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblFunctionName,lblFunctionName,-1,WordWrap
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns/sets a value that determines whether a control expands to fit the text in its Caption."
    WordWrap = lblFunctionName.WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    lblFunctionName.WordWrap() = New_WordWrap
    PropertyChanged "WordWrap"
End Property

'========================================================================
'Control Button Properties
'
Private Sub cmdControlButton_Click(Index As Integer)
    RaiseEvent ButtonClick(Index)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdControlButton(0),cmdControlButton,0,Font
Public Property Get ButtonFont() As Font
Attribute ButtonFont.VB_Description = "Returns a Font object."
    Set ButtonFont = cmdControlButton(0).Font
End Property

Public Property Set ButtonFont(ByVal New_ButtonFont As Font)
    Set cmdControlButton(0).Font = New_ButtonFont
    Set cmdControlButton(1).Font = New_ButtonFont
    Set cmdControlButton(2).Font = New_ButtonFont
    Set cmdControlButton(3).Font = New_ButtonFont
    Set cmdControlButton(4).Font = New_ButtonFont
    Set cmdControlButton(5).Font = New_ButtonFont
    PropertyChanged "ButtonFont"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ButtonHeight() As Long
Attribute ButtonHeight.VB_ProcData.VB_Invoke_Property = "ControlButton"
    ButtonHeight = cmdControlButton(0).Height
End Property

Public Property Let ButtonHeight(ByVal New_ButtonHeight As Long)
    cmdControlButton(0).Height = New_ButtonHeight
    cmdControlButton(1).Height = New_ButtonHeight
    cmdControlButton(2).Height = New_ButtonHeight
    cmdControlButton(3).Height = New_ButtonHeight
    cmdControlButton(4).Height = New_ButtonHeight
    cmdControlButton(5).Height = New_ButtonHeight
    PropertyChanged "ButtonHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ButtonWidth() As Long
Attribute ButtonWidth.VB_ProcData.VB_Invoke_Property = "ControlButton"
    ButtonWidth = cmdControlButton(0).Width
End Property

Public Property Let ButtonWidth(ByVal New_ButtonWidth As Long)
    cmdControlButton(0).Width = New_ButtonWidth
    cmdControlButton(1).Width = New_ButtonWidth
    cmdControlButton(2).Width = New_ButtonWidth
    cmdControlButton(3).Width = New_ButtonWidth
    cmdControlButton(4).Width = New_ButtonWidth
    cmdControlButton(5).Width = New_ButtonWidth
    PropertyChanged "ButtonWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ButtonTop() As Long
Attribute ButtonTop.VB_ProcData.VB_Invoke_Property = "ControlButton"
    ButtonTop = cmdControlButton(0).Top
End Property

Public Property Let ButtonTop(ByVal New_ButtonTop As Long)
    cmdControlButton(0).Top = New_ButtonTop
    cmdControlButton(1).Top = New_ButtonTop
    cmdControlButton(2).Top = New_ButtonTop
    cmdControlButton(3).Top = New_ButtonTop
    cmdControlButton(4).Top = New_ButtonTop
    cmdControlButton(5).Top = New_ButtonTop
    PropertyChanged "ButtonTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdControlButton(0),cmdControlButton,0,Caption
Public Property Get ButtonCaption(Index As Integer) As String
Attribute ButtonCaption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute ButtonCaption.VB_ProcData.VB_Invoke_Property = "ControlButton"
    ButtonCaption(Index) = cmdControlButton(Index).Caption
End Property

Public Property Let ButtonCaption(Index As Integer, ByVal New_ButtonCaption As String)
    cmdControlButton(Index).Caption = New_ButtonCaption
    PropertyChanged "ButtonCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ButtonLeft(Index As Integer) As Long
    ButtonLeft(Index) = cmdControlButton(Index).Left
End Property

Public Property Let ButtonLeft(Index As Integer, ByVal New_ButtonLeft As Long)
    cmdControlButton(Index).Left = New_ButtonLeft
    PropertyChanged "ButtonLeft"
End Property

'
'Control Button Properties (end)
'========================================================================
'
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get HeadingCaption() As String
Attribute HeadingCaption.VB_ProcData.VB_Invoke_Property = "HeadingLabel"
    HeadingCaption = m_HeadingCaption
End Property

Public Property Let HeadingCaption(ByVal New_HeadingCaption As String)
    m_HeadingCaption = New_HeadingCaption
    PropertyChanged "HeadingCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get HeadingFont() As Font
    Set HeadingFont = m_HeadingFont
End Property

Public Property Set HeadingFont(ByVal New_HeadingFont As Font)
    Set m_HeadingFont = New_HeadingFont
    PropertyChanged "HeadingFont"
End Property

Public Property Let QueryRecordset(rs As ADODB.Recordset)
    m_QueryRecordset = rs
    If Not ShowDetail(m_QueryRecordset) Then
    End If
    PropertyChanged "QueryRecordset"
End Property

Public Function ShowDetail(rs As ADODB.Recordset) As Boolean
Dim sGridString As String
Dim iCount As Integer
Dim iCounter As Integer
Dim iNumberOfFields As Integer

On Error GoTo ErrorHandler

ShowDetail = False
sGridString = ""
iNumberOfFields = rs.Fields.Count - 1
    
With fgDataGrid
    .Redraw = False
    .Clear
    .Cols = rs.Fields.Count
    .Rows = 0
    .AddItem sGridString
    
'Load search detail
    For iCount = 1 To rs.RecordCount
        For iCounter = 0 To iNumberOfFields
            .ColAlignment(iCounter) = 1
            sGridString = sGridString & rs.Fields(iCounter).Value & vbTab
            .MergeCol(iCounter) = True     ' Allow merge on Columns 0 thru 3
        Next iCounter
        .AddItem sGridString
        rs.MoveNext
        sGridString = ""
    Next iCount

    .FixedRows = 1
'Load heading row
    For iCounter = 0 To iNumberOfFields
        .TextMatrix(0, iCounter) = rs.Fields(iCounter).Name
    Next iCounter
    
    .Redraw = True
End With

ShowDetail = True
Exit Function

ErrorHandler:
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property


