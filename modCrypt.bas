Attribute VB_Name = "modCrypt"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : Encrypt
' DateTime  : 31-03-2003 18:55 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Encrypts the given text by saving it as a picture
'---------------------------------------------------------------------------------------
Public Function Encrypt(FromText As String, ToPicture As PictureBox) As String
   
  On Error GoTo EncryptError

  ToPicture.ScaleMode = vbPixels 'Sets the picturebox to use pixels
  ToPicture.Cls                  'Clears the picture of previous pixels
  
  'Dim's variables to hold the ASCII-codes of three character
  Dim Char1 As Integer, Char2 As Integer, Char3 As Integer
  Dim Xpos As Long, Ypos As Long 'The current position for the pixel
  
  'Makings sure the lenth of the text are divideable with 3
  FromText = FromText & String(3 - (Len(FromText) Mod 3), Chr(0))
  
  Dim ix As Long 'Counter
  For ix = 1 To Len(FromText) 'For each character in "FromText"
    
    'For each pixel (three char's to make one pixel):
    If ix Mod 3 = 0 Then
      'Extracts the ASCII-codes of each of the three letters
      Char1 = Asc(Mid(FromText, ix - 2, 1))
      Char2 = Asc(Mid(FromText, ix - 1, 1))
      Char3 = Asc(Mid(FromText, ix, 1))
      
      'Xpos = PixelNumber - PreviosPixels      (PixelNumber starts width 0)
      Xpos = ((ix / 3) - 1) - (ToPicture.ScaleWidth * Ypos)
      
      'If the xpos is greater than the picture's width:
      If Xpos > (ToPicture.ScaleWidth - 1) Then
        'Moves to the next line of pixels
        Ypos = Ypos + 1
        'Resets the x-position
        Xpos = 0
        'If the ypos is greater than the picture's height:
        If Ypos > ToPicture.ScaleHeight Then
          Dim LostPixels As Integer
          'Counts the number of lost pixels
          LostPixels = Round((Len(frmMain.rtxtEncrypt.Text) / 3) + 1.5) - (ToPicture.ScaleWidth * Ypos)
          
          MsgBox "The picture cacvas is too small to hold all the pixels." & vbNewLine & _
             LostPixels & " pixels are lost!", vbCritical
          Exit For 'Exits the loop
        End If
      End If
      
      'Draws the current pixel by using the three ASCII-codes as the RGB-values
      ToPicture.PSet (Xpos, Ypos), rgb(Char1, Char2, Char3)
    End If

  Next ix

  'Sets a "Null"-pixel consisting of three Chr(0)'s
  ToPicture.PSet (Len(FromText) / 3, 0), rgb(0, 0, 0)
  Encrypt = "Encryption complete!" 'Returning message
  
  Exit Function
EncryptError:
  Encrypt = "An error occured!" 'Returning message
End Function

'---------------------------------------------------------------------------------------
' Procedure : Decrypt
' DateTime  : 31-03-2003 18:56 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Read the characterinfomation from the pixels in the picture
'---------------------------------------------------------------------------------------
Public Function Decrypt(FromPicture As PictureBox) As String
  FromPicture.ScaleMode = vbPixels
  
  Dim TempText As String, PixelColor As OLE_COLOR, Ypos As Long, Runing As Boolean
  Dim iy As Long, ix As Long 'Conuters
  'Dim's variables to hold the characters of the three ASCII-codes
  Dim Char1 As String, Char2 As String, Char3 As String
  
  'Variable used to exit the loop when encountering a Null-pixel
  Runing = True
  
  'Scanning from the picture's top and down to the buttom
  For iy = 0 To FromPicture.ScaleHeight - 1
    'Scanning from the picture's left and right to the picture's width is reached
    For ix = 0 To FromPicture.ScaleWidth - 1
    
        'Reads the OLE COLOR-value of the current pixel
        PixelColor = FromPicture.Point(ix, Ypos)
        'Extracts the RGB-values from the pixel and converting them to characters
        Char1 = Chr(RedFromRGB(PixelColor))
        Char2 = Chr(GreenFromRGB(PixelColor))
        Char3 = Chr(BlueFromRGB(PixelColor))
        'Saving them in a string
        TempText = TempText & Char1 & Char2 & Char3
        'If a Null-pixel is encounted:
        If Char1 = Chr(0) Or Char2 = Chr(0) Or Char3 = Chr(0) Then
          Runing = False
          Exit For 'Exits inner loop
        End If
        
    Next ix
    Ypos = Ypos + 1 'Move to next line of pixels
    If Runing = False Then Exit For 'Exits the outter loop
  Next iy
  
  'Replacing the Null-pixel width nothing ("")
  TempText = Replace(TempText, Chr(0), "")
  'Returns the decrypted text
  Decrypt = TempText
End Function

'---------------------------------------------------------------------------------------
' Procedure : RedFromRGB
' DateTime  : 31-03-2003 19:29 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Retrives the red value from a RGB-color
'---------------------------------------------------------------------------------------
Public Function RedFromRGB(ByVal rgb As Long) As Integer
  RedFromRGB = &HFF& And rgb
End Function

'---------------------------------------------------------------------------------------
' Procedure : GreenFromRGB
' DateTime  : 31-03-2003 19:30 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Retrives the green value from a RGB-color
'---------------------------------------------------------------------------------------
Public Function GreenFromRGB(ByVal rgb As Long) As Integer
  GreenFromRGB = (&HFF00& And rgb) \ 256
End Function

'---------------------------------------------------------------------------------------
' Procedure : BlueFromRGB
' DateTime  : 31-03-2003 19:30 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Retrives the blue value from a RGB-color
'---------------------------------------------------------------------------------------
Public Function BlueFromRGB(ByVal rgb As Long) As Integer
  BlueFromRGB = (&HFF0000 And rgb) \ 65536
End Function
