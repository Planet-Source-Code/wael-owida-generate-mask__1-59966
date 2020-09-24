VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Just We use for this sample only tow functions from below
' its GetPixel and SetPixel
' BUT THE OTHER FUNCTIONs and what i will add of extra functions
' Because my next update I will update this class To DO
' Many Handle To Bitmap


Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Option Explicit

Public Sub GenerateMaskBitmap(ByRef Bmp As PictureBox, ByRef Bmp2 As PictureBox, ByVal ColorTransparent As Long, ByVal ColorBitmap As Long, ByVal ColorMask As Long)
' For Best mask use type "BMP" Bitmap , u can use jpg and gif but for best use BMP

Dim IW As Long, IH As Long
Dim LW As Long, LH As Long

With Bmp
    .ScaleMode = vbPixels
    .AutoRedraw = True
    .AutoSize = True
    IW = .ScaleWidth
    IH = .ScaleHeight
    .AutoSize = False
    .Width = Bmp2.Width
    .Height = Bmp2.Height
End With
With Bmp2
    .ScaleMode = vbPixels
    .AutoRedraw = True
    .Cls
    .Refresh
Set .Picture = Nothing
End With
For LW = 1 To IW
    For LH = 1 To IH
        If GetPixel(Bmp.hdc, LW, LH) <> ColorTransparent Then
            SetPixel Bmp2.hdc, LW, LH, ColorBitmap
        Else
            SetPixel Bmp2.hdc, LW, LH, ColorMask
        End If
    Next
Next
Bmp2.Picture = Bmp2.Image
End Sub
Public Sub SaveMaskToFile(ByRef Bmp As PictureBox, ByVal SaveToFile As String)
If SaveToFile <> "" Then
    If LCase(Right$(SaveToFile, 3)) = LCase("Bmp") Then
        SavePicture Bmp.Picture, SaveToFile
    End If
If Dir$(SaveToFile) = "" Then Debug.Print "Bad FileName Or Format, File Not Saved"
End If
End Sub
