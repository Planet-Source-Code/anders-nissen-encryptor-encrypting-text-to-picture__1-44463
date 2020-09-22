VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Encryptor"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10830
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   5250
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   8040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjEncryptor.BoxFrame bxfPicture 
      Height          =   1305
      Left            =   180
      TabIndex        =   6
      Top             =   3765
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   2302
      Caption         =   "Encrypted"
      Begin VB.PictureBox picEncrypted 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   268
         TabIndex        =   7
         Top             =   200
         Width           =   4020
      End
      Begin VB.Shape shpPicBorder 
         BorderColor     =   &H00E0E0E0&
         Height          =   135
         Left            =   1560
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblPictureInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1060
         Width           =   3975
      End
   End
   Begin prjEncryptor.BoxFrame bxfText 
      Height          =   2930
      Left            =   180
      TabIndex        =   3
      Top             =   750
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   5159
      Caption         =   "Text To Encrypt"
      Begin RichTextLib.RichTextBox rtxtEncrypt 
         Height          =   2520
         Left            =   120
         TabIndex        =   4
         Top             =   135
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4445
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMain.frx":2824
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblTextInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   3975
      End
   End
   Begin prjEncryptor.BoxFrame bxfOptions 
      Height          =   4335
      Left            =   4560
      TabIndex        =   9
      Top             =   750
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7646
      Caption         =   "Options"
      Begin VB.PictureBox picOptions 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3405
         Index           =   2
         Left            =   4200
         ScaleHeight     =   3375
         ScaleWidth      =   1905
         TabIndex        =   27
         Top             =   420
         Width           =   1935
         Begin VB.CommandButton cmdOptions 
            Caption         =   "OK"
            Enabled         =   0   'False
            Height          =   225
            Index           =   13
            Left            =   1440
            TabIndex        =   37
            Tag             =   "Accepts and applies the chosen width"
            Top             =   1835
            Width           =   495
         End
         Begin VB.TextBox txtPicWidth 
            Enabled         =   0   'False
            Height          =   270
            Left            =   435
            MaxLength       =   3
            TabIndex        =   35
            Tag             =   "Enter a new pixel value for the picture's width"
            Text            =   "width"
            Top             =   1800
            Width           =   495
         End
         Begin VB.CheckBox chkLimitPicWidth 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Limit picturewidth"
            Height          =   225
            Left            =   0
            TabIndex        =   34
            Tag             =   "Limit the width of the picture to a given value"
            Top             =   1560
            Width           =   1935
         End
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Restore Default"
            Height          =   375
            Index           =   11
            Left            =   0
            TabIndex        =   42
            Tag             =   "Restores the settings to the default"
            Top             =   3000
            Width           =   1935
         End
         Begin VB.CheckBox chkFade 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fade in/out"
            Height          =   180
            Left            =   0
            TabIndex        =   30
            Tag             =   "Fades the program at startup and shutdown"
            Top             =   480
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkPictureInfo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show picture info"
            Height          =   180
            Left            =   0
            TabIndex        =   41
            Tag             =   "Show infomation about the picture "
            Top             =   2760
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkTextInfo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show text info"
            Height          =   180
            Left            =   0
            TabIndex        =   40
            Tag             =   "Show infomation about the text in the textbox"
            Top             =   2520
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkOptionsInfo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show options info"
            Height          =   180
            Left            =   0
            TabIndex        =   39
            Tag             =   "Show infomation about the different controls (this)"
            Top             =   2280
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkAutoPaste 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Paste text at startup"
            Height          =   180
            Left            =   0
            TabIndex        =   33
            Tag             =   "Pastes the clipboard to the textbox at startup"
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox chkAniMenus 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Animates option tabs"
            Height          =   180
            Left            =   0
            TabIndex        =   29
            Tag             =   "Sliding effect in the ""options""-tabs"
            Top             =   240
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkAutoScramble 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Auto-scramble"
            Height          =   180
            Left            =   0
            TabIndex        =   32
            Tag             =   "Automatically scrambles the encrypted picture"
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lblPixels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "pixels"
            Enabled         =   0   'False
            Height          =   180
            Left            =   945
            TabIndex        =   36
            Top             =   1830
            Width           =   450
         End
         Begin VB.Label lblSettings 
            BackStyle       =   0  'Transparent
            Caption         =   "Info"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   38
            Top             =   2040
            Width           =   1935
         End
         Begin VB.Label lblSettings 
            BackStyle       =   0  'Transparent
            Caption         =   "Functions"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   31
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label lblSettings 
            BackStyle       =   0  'Transparent
            Caption         =   "Appearance"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.PictureBox picOptions 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3405
         Index           =   1
         Left            =   2160
         ScaleHeight     =   3375
         ScaleWidth      =   1905
         TabIndex        =   19
         Top             =   420
         Width           =   1935
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Unscramble Picture"
            Height          =   375
            Index           =   12
            Left            =   0
            TabIndex        =   26
            Tag             =   "Makes a scrambled picture decryptable"
            Top             =   2880
            Width           =   1935
         End
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Paste Picture"
            Height          =   375
            Index           =   9
            Left            =   0
            TabIndex        =   24
            Tag             =   "Paste the context of the clipboard into the picture"
            Top             =   1920
            Width           =   1935
         End
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Scramble Picture"
            Height          =   375
            Index           =   10
            Left            =   0
            TabIndex        =   25
            Tag             =   "Modifies the picture so it isn't decryptable"
            Top             =   2400
            Width           =   1935
         End
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Copy Picture"
            Height          =   375
            Index           =   8
            Left            =   0
            TabIndex        =   23
            Tag             =   "Copies the context of the picture to the clipboard"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Load Picture"
            Height          =   375
            Index           =   6
            Left            =   0
            TabIndex        =   21
            Tag             =   "Loads an encrypted file for decryption"
            Top             =   480
            Width           =   1935
         End
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Save Picture"
            Height          =   375
            Index           =   7
            Left            =   0
            TabIndex        =   22
            Tag             =   "Saves the encrypted text at a specified destination"
            Top             =   960
            Width           =   1935
         End
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Decrypt Picture"
            Height          =   375
            Index           =   5
            Left            =   0
            TabIndex        =   20
            Tag             =   "Reads the characterinfo from the picture"
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.PictureBox picOptions 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3405
         Index           =   0
         Left            =   120
         ScaleHeight     =   3375
         ScaleWidth      =   1905
         TabIndex        =   13
         Top             =   420
         Width           =   1935
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Paste Text"
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   18
            Tag             =   "Paste the context of the clipboard to the textbox"
            Top             =   1920
            Width           =   1935
         End
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Encrypt Text"
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Tag             =   "Converts the text to a picture"
            Top             =   0
            Width           =   1935
         End
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Copy Text"
            Height          =   375
            Index           =   3
            Left            =   0
            TabIndex        =   17
            Tag             =   "Copies the text from the textbox to the clipboard"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Save Text"
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   16
            Tag             =   "Saves the text at a specified destination"
            Top             =   960
            Width           =   1935
         End
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Load Text"
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   15
            Tag             =   "Loads a file into the textbox for encryption"
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Label lblOptionsInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   3880
         Width           =   1935
      End
      Begin VB.Label lblChooseOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Settings"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   12
         Tag             =   "Customize settings to fit your needs"
         Top             =   170
         Width           =   735
      End
      Begin VB.Label lblChooseOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Picture"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   11
         Tag             =   "Contains the tools for handling the picture"
         Top             =   170
         Width           =   615
      End
      Begin VB.Label lblChooseOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Text"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Tag             =   "Contains the tools for handling the text"
         Top             =   170
         Width           =   615
      End
   End
   Begin VB.Label lblTitel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IcySoft 2003"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   2
      Left            =   4605
      TabIndex        =   2
      Top             =   210
      Width           =   1080
   End
   Begin VB.Label lblTitel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VERSION"
      ForeColor       =   &H00404040&
      Height          =   180
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   225
      Width           =   765
   End
   Begin VB.Image imgUnload 
      Height          =   255
      Left            =   6620
      Top             =   50
      Width           =   255
   End
   Begin VB.Image imgMinimize 
      Height          =   255
      Left            =   6330
      Top             =   50
      Width           =   255
   End
   Begin VB.Image imgFormTop 
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label lblTitel 
      BackStyle       =   0  'Transparent
      Caption         =   "Encryptor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   200
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmMain
' DateTime  : 01-04-2003 23:14 CET
' Author    : Anders Nissen, IcySoft
' Copyright : IcySoft 2003
' Purpose   : The main (and so far only) form.
'---------------------------------------------------------------------------------------

Option Explicit
Dim picWidth As Integer, PicHeight As Integer

Private Sub Form_Load()
  'Sets the form to the size of the skin-picture
  Me.Width = 6930
  Me.Height = 5250
  'Rounds the edges to make to app look smooth :)
  RoundEdges Me
  
  'Saves the picturebox's height and width in static variables for later use
  picWidth = picEncrypted.Width
  PicHeight = picEncrypted.Height
  'Sets the width of the picturebox and applies the fancy border
  SetPicWidth picWidth
  
  'Loads the values of the controls of the "Settings"-page from the registry
  LoadOptions
  
  'Dynamicly sets the version-number
  lblTitel(1).Caption = " v." & App.Major & "." & App.Minor & "." & App.Revision
  
  'Hiding each of the options-pages exept the "Text"-page
  Dim ix As Integer, LowIx As Integer
  LowIx = picOptions.LBound
  
  For ix = LowIx To picOptions.uBound
    picOptions(ix).Visible = False
    picOptions(ix).Move picOptions(LowIx).Left, picOptions(LowIx).Top, _
        picOptions(LowIx).Width, picOptions(LowIx).Height
    picOptions(ix).BorderStyle = 0 'Removes the border
  Next ix
  picOptions(LowIx).Visible = True 'Shows the "Text"-page
  
  bxfOptions.Width = 2175 'Resizes the frame containing the options-pages
   
  Me.Show 'Shows the form
  If chkFade.Value = vbChecked Then FormFadeIn Me 'Fades if enabled
  
  'Auto-pastes the text of the clipboard to the textbox if enabled
  If chkAutoPaste.Value = vbChecked Then rtxtEncrypt.Text = Clipboard.GetText
End Sub

Private Sub bxfOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Removes underline from each of the labels
  Dim ix As Integer
  For ix = lblChooseOption.LBound To lblChooseOption.uBound
    If lblChooseOption(ix).FontUnderline <> 0 Then lblChooseOption(ix).FontUnderline = 0
  Next ix
End Sub

Private Sub chkLimitPicWidth_Click()
  'Enables/disabels the controls linked to chkLimitPicWidth
  txtPicWidth.Enabled = chkLimitPicWidth.Value
  lblPixels.Enabled = chkLimitPicWidth.Value
  cmdOptions(13).Enabled = chkLimitPicWidth.Value
  
  'User un-checked the checkbox:
  If chkLimitPicWidth.Value = vbUnchecked Then
    'Apply default width
    SetPicWidth picWidth
  Else
    'Calls the "OK"-button
    cmdOptions_Click 13
    If txtPicWidth.Visible = True Then txtPicWidth.SetFocus
  End If
End Sub

Private Sub chkOptionsInfo_Click()
  'Shows/hides the infomation-label
  lblOptionsInfo.Visible = chkOptionsInfo.Value
End Sub

Private Sub chkPictureInfo_Click()
  'Shows/hides the infomation-label
  lblPictureInfo.Visible = chkPictureInfo.Value
End Sub

Private Sub chkTextInfo_Click()
  'Shows/hides the infomation-label
  lblTextInfo.Visible = chkTextInfo.Value
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdOptions_Click
' DateTime  : 03-04-2003 16:53 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : This sub handles all the events of the commandboxes in the option-pages
'---------------------------------------------------------------------------------------
Private Sub cmdOptions_Click(Index As Integer)
  On Error GoTo cmdOptions_Click_Error

  Select Case cmdOptions(Index).Caption
    
    '//   TEXT   \\
    Case "Encrypt Text"
      lblPictureInfo.Caption = "Encrypting..."
      lblPictureInfo.Refresh
      'Encrypts the text!
      If isTrueColor Then lblPictureInfo.Caption = Encrypt(rtxtEncrypt.Text, picEncrypted)
      'Scrambles the picture if enabled
      If chkAutoScramble.Value = vbChecked Then Scramble picEncrypted, 2
      
    Case "Load Text"
      Dim TextSourceToLoad As String
      TextSourceToLoad = LoadPath("Text files|*.txt")
      'Loads the chosen file
      If TextSourceToLoad <> "" Then rtxtEncrypt.LoadFile TextSourceToLoad
    
    Case "Save Text"
      Dim TextDestToSave As String
      TextDestToSave = SavePath("Text files|*.txt")
      'Saves the text to the chosen destination
      If TextDestToSave <> "" Then rtxtEncrypt.SaveFile TextDestToSave
      
    Case "Copy Text"
      'Copies the text of the textbox to the clipboard
      Clipboard.SetText rtxtEncrypt.Text
      
    Case "Paste Text"
      'Pastes the text from the clipboard after the text in the textbox
      rtxtEncrypt.Text = rtxtEncrypt.Text & Clipboard.GetText
    
    '// PICTURE  \\
    Case "Decrypt Picture"
      lblTextInfo.Caption = "Decrypting...": lblTextInfo.Refresh
      'Decrypts the picture!
      If isTrueColor Then rtxtEncrypt.Text = Decrypt(picEncrypted)
      lblTextInfo.Caption = "Decryption complete!"
    
    Case "Load Picture"
      Dim PicSourceToLoad As String
      PicSourceToLoad = LoadPath("Bitmap files|*.bmp")
      'Loads the chosen file
      If PicSourceToLoad <> "" Then picEncrypted.Picture = LoadPicture(PicSourceToLoad)
        
    Case "Save Picture"
      Dim PicDestToSave As String
      PicDestToSave = SavePath("Bitmap files|*.bmp")
      'Saves the picture to the chosen destination
      If PicDestToSave <> "" Then SavePicture picEncrypted.Image, PicDestToSave
    
    Case "Copy Picture"
      'Copies the content of the picturebox to the clipboard
      Clipboard.SetData picEncrypted.Image
    
    Case "Paste Picture"
      'Pastes the picture, if any, from the clipboard to the picturebox
      picEncrypted.Picture = Clipboard.GetData
    
    Case "Scramble Picture"
      'Scrambles the picturebox making the picture un-readable (by this app anyway ;) )
      Scramble picEncrypted, 2 'InputBox("Enter a scramblecode between 1-10", , 5)
    
    Case "Unscramble Picture"
      'Un-scrambles a scrambled picture to make it readable
      Scramble picEncrypted, 2, True 'InputBox("Enter a scramblecode between 1-10", , 5), True
    
    '// SETTINGS \\
    
    Case "Restore Default"
      'Sets the default values of the controls in the "Settings"-page
      chkAniMenus.Value = vbChecked
      chkFade.Value = vbChecked
      chkAutoScramble.Value = vbUnchecked
      chkAutoPaste.Value = vbUnchecked
      chkLimitPicWidth.Value = vbUnchecked
      chkLimitPicWidth.Caption = "Limit picturewidth"
      picEncrypted.Width = picWidth
      chkOptionsInfo.Value = vbChecked
      chkTextInfo.Value = vbChecked
      chkPictureInfo.Value = vbChecked
    
    Case "OK"
      Dim WidthLimit As Integer
      'Making sure the text is numeric and in twips
      WidthLimit = Val(txtPicWidth.Text) * Screen.TwipsPerPixelX 'In twips
    
      'Width is too small, too large or mistyped:
      If WidthLimit < 75 Or WidthLimit > picWidth Then 'All in twips
        MsgBox "Width outside limit or mistyped statement." & vbNewLine & _
          "Enter a value between 5-" & picWidth / Screen.TwipsPerPixelX, vbExclamation
        'Sets the text to the starting width of the picturebox
        txtPicWidth.Text = picWidth / Screen.TwipsPerPixelX
        'Sets the picturebox's width to the starting width of the picturebox
        SetPicWidth picWidth 'In twips
      Else
        'Given value is OK - applies the value to picturebox
        SetPicWidth WidthLimit 'In twips
      End If
      
      'Set focus if visible
      If txtPicWidth.Visible = True Then txtPicWidth.SetFocus
    
    Case Else
    '...hmmm, just has to be here!
    
  End Select

  'Error - ie. wrong datatype pasted into the picturebox ect.
  On Error GoTo 0
  Exit Sub
cmdOptions_Click_Error:
  MsgBox "An error occured!" & vbNewLine & _
    "Error " & Err.Number & " (" & Err.Description & ")", vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Fades the form if Fading is enabled
  If chkFade.Value = vbChecked Then FormFadeOut Me
  SaveOptions 'Saves the options in the registry
  End 'Closes the program
End Sub

Private Sub imgFormTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Moves the form
  MoveForm Me
End Sub

Private Sub imgMinimize_Click()
  'Minimizing the form
  Me.WindowState = 1
End Sub

Private Sub imgUnload_Click()
  'Referers to the "Form_Unload"-sub
  Unload Me
End Sub

Private Sub lblChooseOption_Click(Index As Integer)
  'If the page is already active don't do anything
  If lblChooseOption(Index).ForeColor = vbBlack Then Exit Sub
  
  'Applying black forecolor to the current label and light-gray to the others
  Dim ix As Integer
  For ix = lblChooseOption.LBound To lblChooseOption.uBound
     lblChooseOption(ix).ForeColor = &HC0C0C0
     picOptions(ix).Visible = False
  Next ix
  lblChooseOption(Index).ForeColor = vbBlack
  lblChooseOption(Index).Refresh
  picOptions(Index).Visible = True 'Showing the selected option-page
  
  'Slides the picturebox if Animation of Optiontabs is enabled
  If chkAniMenus.Value = vbChecked Then
    For ix = 0 To picOptions(Index).Height Step picOptions(Index).Height / 35
      picOptions(Index).Height = ix
      picOptions(Index).Refresh
    Next ix
  End If
End Sub

Private Sub lblChooseOption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Applying underline to the current label and removing it from the others
  Dim ix As Integer
  For ix = lblChooseOption.LBound To lblChooseOption.uBound
    If lblChooseOption(ix).FontUnderline <> 0 Then lblChooseOption(ix).FontUnderline = 0
  Next ix
  If lblChooseOption(Index).FontUnderline <> 1 Then lblChooseOption(Index).FontUnderline = 1
  'Writing the tag-info to lblOptionsInfo
  lblOptionsInfo.Caption = lblChooseOption(Index).Tag
End Sub

Private Sub lblTitel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Selfexplaining - moves the form using API (see the function under "modAPI")
  MoveForm Me
End Sub

Private Sub rtxtEncrypt_Change()
  'Info about the text in the textbox
  lblTextInfo.Caption = "Characters: " & Len(rtxtEncrypt.Text) & " Pixels: " & _
    Round((Len(rtxtEncrypt.Text) / 3) + 1.5) 'Length of text in pixels
End Sub

Private Sub txtPicWidth_GotFocus()
  'Selects the text, if any, in the textbox when got focus
  txtPicWidth.SelStart = 0
  txtPicWidth.SelLength = Len(txtPicWidth.Text)
  lblOptionsInfo.Caption = txtPicWidth.Tag
End Sub

Private Sub picEncrypted_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Extracts the color- and char-values from the current (x,y) of the picture
  Dim PixelFarve As OLE_COLOR, Tegn1, Tegn2, Tegn3, Colors
  PixelFarve = picEncrypted.Point(X, Y)
  Tegn1 = Chr(RedFromRGB(PixelFarve))
  Tegn2 = Chr(GreenFromRGB(PixelFarve))
  Tegn3 = Chr(BlueFromRGB(PixelFarve))
  Colors = RedFromRGB(PixelFarve) & "," & GreenFromRGB(PixelFarve) & "," & BlueFromRGB(PixelFarve)
  'Shows the info in the label below the picture
  lblPictureInfo.Caption = "(" & X & "," & Y & ") RGB: " & Colors & " Characters: " & Tegn1 & Tegn2 & Tegn3
End Sub

Private Sub chkFade_KeyDown(KeyCode As Integer, Shift As Integer)
  'If the function is grayed (but not "enabled" to still allow the info the by shown)
  If chkFade.Value = vbGrayed Then
    'Explain why the function is unusable
    MsgBox "This function is not functional in your Operation System", vbInformation
    'Make sure the checkbox isn't changed
    chkFade.Value = vbGrayed
  End If
End Sub

Private Sub chkFade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Refers to the chkFade_keyDown-sub
  chkFade_KeyDown 0, 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdOptions_MouseMove
' DateTime  : 03-04-2003 16:32 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : The following sub's are just for applying info about the controls to lblOptionsInfo
'---------------------------------------------------------------------------------------
Private Sub cmdOptions_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblOptionsInfo.Caption = cmdOptions(Index).Tag
End Sub

Private Sub chkAniMenus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblOptionsInfo.Caption = chkAniMenus.Tag
End Sub

Private Sub txtPicWidth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblOptionsInfo.Caption = txtPicWidth.Tag
End Sub

Private Sub chkTextInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblOptionsInfo.Caption = chkTextInfo.Tag
End Sub

Private Sub chkAutoPaste_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblOptionsInfo.Caption = chkAutoPaste.Tag
End Sub

Private Sub chkPictureInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblOptionsInfo.Caption = chkPictureInfo.Tag
End Sub

Private Sub chkAutoScramble_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblOptionsInfo.Caption = chkAutoScramble.Tag
End Sub

Private Sub chkFade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblOptionsInfo.Caption = chkFade.Tag
End Sub

Private Sub chkLimitPicWidth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblOptionsInfo.Caption = chkLimitPicWidth.Tag
End Sub

Private Sub chkOptionsInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblOptionsInfo.Caption = chkOptionsInfo.Tag
End Sub
