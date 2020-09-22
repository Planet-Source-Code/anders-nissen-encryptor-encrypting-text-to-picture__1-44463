Attribute VB_Name = "modMiscFunctions"
'---------------------------------------------------------------------------------------
' Module    : modMiscFunctions
' DateTime  : 01-04-2003 23:11 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : This module holds most of the non-API miscellaneous function used
'---------------------------------------------------------------------------------------
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : SavePath
' DateTime  : 01-04-2003 10:00 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Finding the destination of source for saving the text/picture
'---------------------------------------------------------------------------------------
Public Function SavePath(Optional FileFilter As String = "") As String
  
      If FileFilter <> "" Then FileFilter = FileFilter & "|"
      
      frmMain.CD.Filter = FileFilter & "All files|*.*"
      frmMain.CD.ShowSave
      If frmMain.CD.FileName <> "" Then
        SavePath = frmMain.CD.FileName
      Else
        MsgBox "No destination selected", vbExclamation
      End If
  

End Function

'---------------------------------------------------------------------------------------
' Procedure : LoadPath
' DateTime  : 01-04-2003 10:03 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Finding the path of the source to load
'---------------------------------------------------------------------------------------
Public Function LoadPath(Optional FileFilter As String = "") As String

  If FileFilter <> "" Then FileFilter = FileFilter & "|"
  
  frmMain.CD.Filter = FileFilter & "All files|*.*"
  frmMain.CD.ShowOpen
  If frmMain.CD.FileName <> "" Then
    LoadPath = frmMain.CD.FileName
  Else
    MsgBox "No file selected", vbExclamation
  End If
  
End Function

'---------------------------------------------------------------------------------------
' Procedure : Scramble
' DateTime  : 01-04-2003 13:54 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Scrambles/Unscrambles the picture to make it not-decryptable/decryptable
'---------------------------------------------------------------------------------------
Public Function Scramble(PicBox As PictureBox, ScrambleValue As Integer, Optional Unscramble As Boolean = False) As Boolean
  If ScrambleValue < 0 Or ScrambleValue > 10 Then
    MsgBox "Scramble value must be between 0-10", vbExclamation
    Exit Function
  End If
  
  PicBox.ScaleMode = vbPixels
  Dim iy As Integer, ix As Integer
  For iy = 0 To PicBox.ScaleHeight
    For ix = 0 To PicBox.ScaleWidth
      If Unscramble = False Then
        PicBox.PSet (ix, iy), PicBox.Point(ix, iy) * ScrambleValue
      Else
        PicBox.PSet (ix, iy), PicBox.Point(ix, iy) / ScrambleValue
      End If
    Next ix
  Next iy

End Function

'---------------------------------------------------------------------------------------
' Procedure : SaveOptions
' DateTime  : 01-04-2003 22:57 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Saves the values of the controls in the options-page in the registry database
'---------------------------------------------------------------------------------------
Public Function SaveOptions()
  SaveInReg "AniMenus", frmMain.chkAniMenus.Value
  SaveInReg "Fade", frmMain.chkFade.Value
  SaveInReg "AutoScramble", frmMain.chkAutoScramble.Value
  SaveInReg "AutoPaste", frmMain.chkAutoPaste.Value
  SaveInReg "LimitPic", frmMain.chkLimitPicWidth.Value
  SaveInReg "LimitPicWidth", Val(frmMain.txtPicWidth.Text)
  SaveInReg "OptionsInfo", frmMain.chkOptionsInfo.Value
  SaveInReg "TextInfo", frmMain.chkTextInfo.Value
  SaveInReg "PictureInfo", frmMain.chkPictureInfo.Value
End Function

'---------------------------------------------------------------------------------------
' Procedure : SaveInReg
' DateTime  : 01-04-2003 23:00 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Handles the registry-save to lighten the code
'---------------------------------------------------------------------------------------
Private Function SaveInReg(Key As String, Value As Variant)
  SaveSetting "Encryptor", "Options", Key, Value
End Function

'---------------------------------------------------------------------------------------
' Procedure : LoadOptions
' DateTime  : 01-04-2003 22:58 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Applies the values of the registry database to the controls in the options-page
'---------------------------------------------------------------------------------------
Public Function LoadOptions()
  frmMain.chkAniMenus.Value = LoadFromReg("AniMenus")
  frmMain.chkFade.Value = LoadFromReg("Fade")
  frmMain.chkAutoScramble.Value = LoadFromReg("AutoScramble", 0)
  frmMain.chkAutoPaste.Value = LoadFromReg("AutoPaste", 0)
  frmMain.txtPicWidth.Text = LoadFromReg("LimitPicWidth", 4035 / 15)
  'Must set value of txtPicWidth before limiting the picture
  frmMain.chkLimitPicWidth.Value = LoadFromReg("LimitPic", 0)
  frmMain.chkOptionsInfo.Value = LoadFromReg("OptionsInfo")
  frmMain.chkTextInfo.Value = LoadFromReg("TextInfo")
  frmMain.chkPictureInfo.Value = LoadFromReg("PictureInfo")
End Function

'---------------------------------------------------------------------------------------
' Procedure : LoadFromReg
' DateTime  : 01-04-2003 22:59 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Handles the registry-load to lighten the code
'---------------------------------------------------------------------------------------
Private Function LoadFromReg(Key As String, Optional Default As Variant = 1) As Variant
  LoadFromReg = GetSetting("Encryptor", "Options", Key, Default)
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetPicWidth
' DateTime  : 02-04-2003 20:50 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Set the width of picEncrypted and moves shpPicBorder
'---------------------------------------------------------------------------------------
Public Sub SetPicWidth(picWidth As Integer)
  With frmMain.picEncrypted
    .Width = picWidth
    frmMain.shpPicBorder.Move .Left - 15, .Top - 15, .Width + 30, .Height + 30
  End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : isTrueColor
' DateTime  : 04-04-2003 20:16 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Checks the color depth of the desktop and returns "true" if it's 32bit
'---------------------------------------------------------------------------------------
Public Function isTrueColor() As Boolean

  Dim ColDepth As Integer, TempResult As Boolean
  ColDepth = modAPI.ColorDepth() 'Extracts the color depth (4,8,16 or 32)
  
  Select Case ColDepth
    Case 4, 8, 16 'If settings isn't true color promt the user
      MsgBox "Your display settings is configured to a color depth of " & ColDepth & " bits." & _
        "To be able to succesfully encrypt and decrypt you must use true-color (32bits)", vbExclamation
      
    Case 32 'Color depth is true color (32bits). Function returns "true"
      TempResult = True
      
    Case Else 'Display setting other that 4,8,16 or 32 bits color depth (propperly not existing)
      MsgBox "Error reading display settings!", vbCritical
  End Select
  
  If TempResult = False Then 'If  true color isn't used
    'Promt the user to change to display settings to 32 bits
    If MsgBox("Do you want to manually configure your display settings to " & _
      "match true color now ?", vbInformation + vbYesNoCancel) = vbYes Then
      'If "Yes" then activates the "Display Settings"-page of the "Monitor Settings"
      Shell "rundll32.exe shell32.dll, Control_RunDLL desk.cpl, ,3", 1
    Else
      'User choose "No" or "Cancel". Cannot succesfully encrypt and decrypt
      MsgBox "You're currently not able to encrypt or decrypt succesfully!", vbExclamation
    End If
  End If
  
  'Returns the result
  isTrueColor = TempResult
End Function
