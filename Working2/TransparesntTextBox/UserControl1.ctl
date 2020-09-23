VERSION 5.00
Begin VB.UserControl TransTextBox 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   ScaleHeight     =   3855
   ScaleWidth      =   4845
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Text1"
      ForeColor       =   &H00000000&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "TransTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=======================================================================================================
' Developer Name:   Paresh Chandra
' Date Created  :   18-Jan-2000
' Control Name  :   TransTextBox
'
' Any modification made to this Code please Add to the modification History and keep this documentation
' When Distributing your software as source Code.
'
' Limitation's
'   There are limitation to this control. First whatever limitations a Label has, this control will also
'   have it. Since the displaying of the text is in a label. It Does not show the Blinking Cursor. So
'   times you will not know if this control has got the focus or not.
'   If there is someone who know of a better solution, please Mail me at c_paresh@email.com. Comments
'   will be Appreciated. There is No Scrolling Bar Implemented.  Didn't Get Time.
'
' Why?
'   I wrote this because I wanted to have text box with a textured background. I looked around but could
'   not find any. I saw the program WinVi. I liked it how you can really customize how text entry looks.
'   Plus if you're writing a program, using Odd Shaped Forms and nice textured Backgrounds, You probably
'   would like to have text entry done that same way.
'   I use this only for Small amount of Text Entry. So I dont have any problem with any limitation of
'   the number of characters.
'
' Future Addition's
'   Scroll bars
'   Mouse Scroll (Click and Drag)
'   Tile Background Bitmap. (Currently Not Working)
'   Custom Property Page for Customizing
'   Password Character
'
'
'=======================================================================================================
' MODIFICATION HISTORY
'-------------------------------------------------------------------------------------------------------
' Date              Developer               Version                     Description
'=======================================================================================================
'
'
' 18-Jan-2000       Paresh                  1.0.0                       Wrote It!
'=======================================================================================================
Option Explicit

'=======================================================================================================
' Windows 32Bit API calls Declared
'=======================================================================================================

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

'=======================================================================================================
' The Variables
'=======================================================================================================

'Default Property Values
Const m_def_Text = 0
Const m_def_TileBitmap = 0
Const m_def_EnableDrag = 0
Const m_def_Appearance = 0

'Property Variables
Dim m_Text          As Variant
Dim m_TileBitmap    As Boolean
Dim m_Picture       As Object
Dim m_EnableDrag    As Boolean
Dim m_Appearance    As Integer


'=======================================================================================================
' Events to Be Mapped
'=======================================================================================================

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Change()

'=======================================================================================================
' Public Enumerations (Long Integer Constants)
'=======================================================================================================

'The Border Style
Public Enum BorderStyle
    None = 0
    Fixed_Single = 1
End Enum

' Back Style
Public Enum BackStyle
    Transparent = 0
    Opaque = 1
End Enum

'Text Alignment
Public Enum Alignment
    Left_Justify = 0
    Right_Justify = 1
    Center = 2
End Enum





Private Sub dragMe()
    If m_def_EnableDrag Then
'        ReleaseCapture
'        Call SendMessage(UserControl.hwnd, &HA1, 2, 0&)
'

    End If
End Sub



Private Sub TileBackground()
'Not Implemented, due to the image being stretched not tiled.

' Dim HorizantalLimit As Integer
'    Dim VerticalLimit As Integer
'    Dim ii As Integer
'    Dim jj As Integer
'
'    HorizantalLimit = Width / UserControl.Width
'    VerticalLimit = Height / UserControl.Height
'    If m_Picture Is Nothing Then
'        Set m_Picture = UserControl.Picture
'
'    End If
'    UserControl.AutoRedraw = True
'   ' m_Picture
'
'
'    For ii = 0 To VerticalLimit - 1
'        For jj = 0 To HorizantalLimit - 1
'            UserControl.PaintPicture m_Picture, jj * UserControl.Width, ii * UserControl.Height, UserControl.Width, UserControl.Height
'        Next
'    Next
End Sub

Private Sub Label1_Click()
    RaiseEvent Click

End Sub

Private Sub UserControl_GotFocus()
Label1.BorderStyle = 1
End Sub

Private Sub UserControl_Initialize()
Label1.Top = 0
Label1.Left = 0
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
If KeyAscii = 13 Then
    Label1 = Label1.Caption & vbCrLf
ElseIf KeyAscii = 8 Then
    If Not Len(Label1.Caption) < 1 Then
        Label1 = Left$(Label1.Caption, (Len(Label1.Caption)) - 1)
    End If

Else
    Label1 = Label1.Caption & Chr$(KeyAscii)
End If
End Sub


Private Sub UserControl_LostFocus()
Label1.BorderStyle = 0
End Sub

Private Sub MoveTheObject(ff As Variant, xx, yy, xButton)
    Static oldx, oldy, mf
    Dim moveleft, movetop
    'Dim ff As Object
    
    
    
    moveleft = ff.Left + xx - oldx
    movetop = ff.Top + yy - oldy
    If xButton = vbLeftButton Then
        If mf = 0 Then
            ff.Move moveleft, movetop
            ff.Refresh
            mf = 1
        Else
            mf = 0
        End If
    End If
    oldx = xx
    oldy = yy
End Sub



Private Sub UserControl_Resize()
Label1.Width = UserControl.Width
Label1.Height = UserControl.Height
If m_TileBitmap Then
    TileBackground
Else
    UserControl.AutoRedraw = False
End If
End Sub


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
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As BackStyle
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
    
       'When You Change the Backstyle. Make Sure you Assign a
    'Picture to the picture Property
    'Same Picture to  MaskePicture Property
    'and Set the Mask Color.
    '
    'This will Allow the Text box to Be Skinned (aka shaped).
End Property

Public Property Let BackStyle(ByVal New_BackStyle As BackStyle)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
    
    'When You Change the Backstyle. Make Sure you Assign a
    'Picture to the picture Property
    'Same Picture to  MaskePicture Property
    'and Set the Mask Color.
    '
    'This will Allow the Text box to Be Skinned (aka shaped).
    
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyle)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Label1.Refresh
End Sub

Private Sub Label1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)

   dragMe

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
   ' Call MoveTheObject(UserControl, X, Y, Button)
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Label1_Change()
    RaiseEvent Change
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MaskColor
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MaskPicture
Public Property Get MaskPicture() As Picture
Attribute MaskPicture.VB_Description = "Returns/sets the picture that specifies the clickable/drawable area of the control when BackStyle is 0 (transparent)."
    Set MaskPicture = UserControl.MaskPicture
End Property

Public Property Set MaskPicture(ByVal New_MaskPicture As Picture)
    Set UserControl.MaskPicture = New_MaskPicture
    PropertyChanged "MaskPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    Set m_Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,WordWrap
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns/sets a value that determines whether a control expands to fit the text in its Caption."
    WordWrap = Label1.WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    Label1.WordWrap() = New_WordWrap
    PropertyChanged "WordWrap"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Text() As Variant
Attribute Text.VB_Description = "Set/Gets the Text of the Transparent Text Box"
    Text = Label1.Caption
End Property

Public Property Let Text(ByVal New_Text As Variant)
    Label1.Caption = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get TileBitmap() As Boolean
Attribute TileBitmap.VB_Description = "To Tile the Picture or Not"
    TileBitmap = m_TileBitmap
End Property

Public Property Let TileBitmap(ByVal New_TileBitmap As Boolean)
    m_TileBitmap = New_TileBitmap
    PropertyChanged "TileBitmap"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Text = m_def_Text
    m_TileBitmap = m_def_TileBitmap
    m_Appearance = m_def_Appearance
    m_EnableDrag = m_def_EnableDrag
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", -2147483633)
    Set MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Label1.WordWrap = PropBag.ReadProperty("WordWrap", False)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_TileBitmap = PropBag.ReadProperty("TileBitmap", m_def_TileBitmap)
    Label1.Alignment = PropBag.ReadProperty("Alignment", 0)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    Label1.AutoSize = PropBag.ReadProperty("AutoSize", False)
    m_EnableDrag = PropBag.ReadProperty("EnableDrag", m_def_EnableDrag)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, -2147483633)
    Call PropBag.WriteProperty("MaskPicture", MaskPicture, Nothing)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("WordWrap", Label1.WordWrap, False)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("TileBitmap", m_TileBitmap, m_def_TileBitmap)
    Call PropBag.WriteProperty("Alignment", Label1.Alignment, 0)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("AutoSize", Label1.AutoSize, False)
    Call PropBag.WriteProperty("EnableDrag", m_EnableDrag, m_def_EnableDrag)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Alignment
Public Property Get Alignment() As Alignment
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = Label1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Alignment)
    Label1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,AutoSize
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
    AutoSize = Label1.AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    Label1.AutoSize() = New_AutoSize
    PropertyChanged "AutoSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get EnableDrag() As Boolean
    EnableDrag = m_EnableDrag
End Property

Public Property Let EnableDrag(ByVal New_EnableDrag As Boolean)
    m_EnableDrag = New_EnableDrag
    PropertyChanged "EnableDrag"
End Property

