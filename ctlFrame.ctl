VERSION 5.00
Begin VB.UserControl BoxFrame 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2550
   ControlContainer=   -1  'True
   ScaleHeight     =   1950
   ScaleWidth      =   2550
   Begin VB.Label lblTitel 
      BackStyle       =   0  'Transparent
      Caption         =   "This is the title!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   -50
      Width           =   1455
   End
   Begin VB.Image imgSide 
      Height          =   75
      Index           =   2
      Left            =   3120
      Picture         =   "ctlFrame.ctx":0000
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   45
   End
   Begin VB.Image imgSide 
      Height          =   75
      Index           =   1
      Left            =   1800
      Picture         =   "ctlFrame.ctx":0082
      Stretch         =   -1  'True
      Top             =   360
      Width           =   60
   End
   Begin VB.Image imgSide 
      Height          =   75
      Index           =   0
      Left            =   2040
      Picture         =   "ctlFrame.ctx":013C
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   45
   End
   Begin VB.Image imgSide 
      Height          =   75
      Index           =   3
      Left            =   1560
      Picture         =   "ctlFrame.ctx":01AE
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   45
   End
   Begin VB.Image imgBorder 
      Height          =   165
      Index           =   2
      Left            =   240
      Picture         =   "ctlFrame.ctx":022C
      Top             =   960
      Width           =   165
   End
   Begin VB.Image imgBorder 
      Height          =   165
      Index           =   3
      Left            =   1080
      Picture         =   "ctlFrame.ctx":03FA
      Top             =   960
      Width           =   165
   End
   Begin VB.Image imgBorder 
      Height          =   165
      Index           =   1
      Left            =   1080
      Picture         =   "ctlFrame.ctx":05C8
      Top             =   240
      Width           =   165
   End
   Begin VB.Image imgBorder 
      Height          =   165
      Index           =   0
      Left            =   240
      Picture         =   "ctlFrame.ctx":0796
      Top             =   240
      Width           =   165
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This is the title!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   175
      TabIndex        =   1
      Top             =   -15
      Width           =   1455
   End
End
Attribute VB_Name = "BoxFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : BoxFrame
' DateTime  : 03-04-2003 16:16 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : A control that functions as a container frame
'---------------------------------------------------------------------------------------
Option Explicit

Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private m_sCaption As String

Public Property Get Caption() As String

  Caption = m_sCaption

End Property

Public Property Let Caption(ByVal sCaption As String)

  m_sCaption = sCaption
  
  lblTitel.Caption = sCaption
  Label1.Caption = sCaption
  
  Call UserControl.PropertyChanged("Caption")

End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  lblTitel.Caption = PropBag.ReadProperty("Caption", "Titel")
  Label1.Caption = PropBag.ReadProperty("Caption", "Titel")
  m_sCaption = PropBag.ReadProperty("Caption", "Titel")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Caption", m_sCaption, "Titel")
End Sub

Private Sub UserControl_Resize()
  'Resizing all the images to create the border of a box
  imgBorder(0).Left = 0
  imgBorder(0).Top = 0
  
  imgBorder(1).Left = UserControl.Width - imgBorder(1).Width
  imgBorder(1).Top = 0
  
  imgBorder(2).Left = 0
  imgBorder(2).Top = UserControl.Height - imgBorder(2).Height
  
  imgBorder(3).Left = UserControl.Width - imgBorder(1).Width
  imgBorder(3).Top = UserControl.Height - imgBorder(2).Height
  
  imgSide(0).Left = imgBorder(0).Width
  imgSide(0).Top = 0
  imgSide(0).Width = UserControl.Width - imgBorder(0).Width - imgBorder(1).Width
  
  imgSide(1).Left = UserControl.Width - imgSide(1).Width
  imgSide(1).Top = imgBorder(1).Height
  imgSide(1).Height = UserControl.Height - imgBorder(1).Width - imgBorder(3).Width
  
  imgSide(2).Left = 0
  imgSide(2).Top = imgBorder(2).Height
  imgSide(2).Height = UserControl.Height - imgBorder(0).Width - imgBorder(2).Width
  
  imgSide(3).Left = imgBorder(2).Width
  imgSide(3).Top = UserControl.Height - imgSide(3).Height
  imgSide(3).Width = UserControl.Width - imgBorder(2).Width - imgBorder(3).Width
End Sub
