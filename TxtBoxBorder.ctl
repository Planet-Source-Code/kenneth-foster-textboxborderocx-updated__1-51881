VERSION 5.00
Begin VB.UserControl TxtBoxBorder 
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   ControlContainer=   -1  'True
   ScaleHeight     =   645
   ScaleWidth      =   615
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   135
      Width           =   345
   End
End
Attribute VB_Name = "TxtBoxBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************
'*    txtBoxBorder OCX                      *
'*    by Ken Foster                         *
'*        2004                              *
'*  Freeware- use anyway you like           *
'********************************************
Option Explicit

Dim FC As Long
Dim NFC As Long

Public Property Get FocusColor() As OLE_COLOR
    FocusColor = FC
End Property

Public Property Let FocusColor(ByVal NewFocusColor As OLE_COLOR)
    FC = NewFocusColor
    UserControl.Line (0, 0)-((UserControl.ScaleLeft + UserControl.ScaleWidth) - 10, (UserControl.ScaleTop + UserControl.ScaleHeight) - 10), FC, B
    PropertyChanged "FocusColor"
End Property

Public Property Get NonFocusColor() As OLE_COLOR
    NonFocusColor = NFC
End Property

Public Property Let NonFocusColor(ByVal NewNonFocusColor As OLE_COLOR)
    NFC = NewNonFocusColor
    UserControl.Line (0, 0)-((UserControl.ScaleLeft + UserControl.ScaleWidth) - 10, (UserControl.ScaleTop + UserControl.ScaleHeight) - 10), NFC, B
    PropertyChanged "NonFocusColor"
End Property

Public Property Get Locked() As Boolean
    Locked = Text1.Locked
End Property

Public Property Let Locked(ByVal NewLocked As Boolean)
    Text1.Locked = NewLocked
    PropertyChanged "Locked"
End Property

Public Property Get Text() As String
    Text = Text1.Text
End Property

Public Property Let Text(ByVal NewText As String)
    Text1.Text = NewText
    PropertyChanged "Text"
End Property

Private Sub Text1_GotFocus()
    UserControl.Line (0, 0)-((UserControl.ScaleLeft + UserControl.ScaleWidth) - 10, (UserControl.ScaleTop + UserControl.ScaleHeight) - 10), FC, B
End Sub

Private Sub Text1_LostFocus()
    UserControl.Line (0, 0)-((UserControl.ScaleLeft + UserControl.ScaleWidth) - 10, (UserControl.ScaleTop + UserControl.ScaleHeight) - 10), NFC, B
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    UserControl.Line (0, 0)-((UserControl.ScaleLeft + UserControl.ScaleWidth) - 10, (UserControl.ScaleTop + UserControl.ScaleHeight) - 10), FC, B
End Sub

Private Sub UserControl_InitProperties()
    FocusColor = vbBlue
    NonFocusColor = vbRed
    
End Sub

Private Sub UserControl_Paint()

    UserControl.Line (0, 0)-((UserControl.ScaleLeft + UserControl.ScaleWidth) - 10, (UserControl.ScaleTop + UserControl.ScaleHeight) - 10), NFC, B
    Text1.Top = UserControl.ScaleTop + 10
    Text1.Left = UserControl.ScaleLeft + 10
    Text1.Width = UserControl.ScaleWidth - 25
    Text1.Height = UserControl.ScaleHeight - 25

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    FC = PropBag.ReadProperty("FocusColor", vbBlue)
    NFC = PropBag.ReadProperty("NonFocusColor", vbRed)
    Text1.Text = PropBag.ReadProperty("Text", "Text")
    Text1.Locked = PropBag.ReadProperty("Locked", False)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Text1.FontSize = PropBag.ReadProperty("FontSize", 8.25)
    Text1.FontBold = PropBag.ReadProperty("FontBold", 0)
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    PropBag.WriteProperty "FocusColor", FC, vbBlue
    PropBag.WriteProperty "NonFocusColor", NFC, vbRed
    PropBag.WriteProperty "Locked", Text1.Locked, False
    PropBag.WriteProperty "Text", Text1.Text, "Text"
    Call PropBag.WriteProperty("FontSize", Text1.FontSize, 0)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", Text1.FontBold, 0)
    
End Sub
Private Sub UserControl_Show()
    
    On Error GoTo ResizeErr
    
    If UserControl.ContainedControls.Count > 0 Then
    
        With UserControl.ContainedControls.Item(0)
            .Top = 10
            .Left = 10
            .Height = UserControl.Height - 25
            .Width = UserControl.Width - 25
        End With

    End If
   
ResizeErr:
    Exit Sub

End Sub

Public Property Get Font() As Font
    Set Font = Text1.Font
End Property
    
Public Property Set Font(ByVal New_Font As Font)
    Set Text1.Font = New_Font
    PropertyChanged "Font"
End Property
    
Public Property Get FontSize() As Single
    FontSize = Text1.FontSize
End Property
    
Public Property Let FontSize(ByVal New_FontSize As Single)
    Text1.FontSize = New_FontSize
    PropertyChanged "FontSize"
End Property
    
Public Property Get FontBold() As Boolean
    FontBold = Text1.FontBold
End Property
    
Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Text1.FontBold = New_FontBold
    PropertyChanged "FontBold"
End Property
