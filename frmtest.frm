VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate Mask"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000C0C0&
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   375
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   8760
      ScaleHeight     =   495
      ScaleWidth      =   2775
      TabIndex        =   18
      Top             =   7080
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000C0C0&
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000C0C0&
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000C0C0&
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5640
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   2520
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "JPG|*.jpg|BMP|*bmp|GIF|*.gif"
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   14
      Top             =   7560
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0FF&
      Height          =   4095
      Left            =   360
      ScaleHeight     =   4035
      ScaleWidth      =   4875
      TabIndex        =   10
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00004080&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "click for Exit"
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "click if u want to save mask in to file"
      Top             =   6600
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0FF&
      Height          =   4095
      Left            =   6600
      ScaleHeight     =   4035
      ScaleWidth      =   4875
      TabIndex        =   7
      Top             =   240
      Width           =   4935
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   6
      ToolTipText     =   "Color to The Mask  , pixel  become for a transparent color"
      Top             =   6600
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   5
      ToolTipText     =   "Color to real Bitmap Pixels"
      Top             =   6120
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   1
      ToolTipText     =   "Color To Transparent"
      Top             =   5640
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Gen Mask"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click to Generate the Mask"
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   6600
      TabIndex        =   13
      Top             =   7642
      Width           =   1140
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Picture"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   6600
      TabIndex        =   11
      Top             =   7162
      Width           =   990
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   2535
      Left            =   0
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   2055
      Left            =   240
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2295
      Left            =   120
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MaskColor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   6600
      TabIndex        =   4
      Top             =   6682
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BitmapColor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   6600
      TabIndex        =   3
      Top             =   6202
      Width           =   1740
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TransColor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   6600
      TabIndex        =   2
      Top             =   5722
      Width           =   1590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINT
    XX As Integer
    YY As Integer
End Type
Private sFile As String
Private Bmp As New clsHBitmap
Private ColorTrans As Long
Private ColorBmp As Long
Private ColorMask As Long
Dim PP As POINT

Option Explicit

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
     MsgBox "Error In Data Entery , please check ur Data Entery", vbExclamation, "Generate Mask"
     Exit Sub
End If
 
If ColorTrans < 0 Then
    MsgBox "Error Color Value In Transparent Color", vbExclamation, "Generate Mask"
    Exit Sub
End If
If ColorBmp < 0 Then
    MsgBox "Error Color Value In Bitmap Color", vbExclamation, "Generate Mask"
    Exit Sub
End If
If ColorMask < 0 Then
    MsgBox "Error Color Value In Mask Color", vbExclamation, "Generate Mask"
    Exit Sub
End If

If sFile = "" Or Picture1.Picture = 0 Then
    MsgBox "Error In Data Entery , please check ur Data Entery", vbExclamation, "Generate Mask"
    Exit Sub
End If

Bmp.GenerateMaskBitmap Picture1, Picture2, ColorTrans, ColorBmp, ColorMask

End Sub

Private Sub Command2_Click()
' this save the mask with the background of the picture2
' if u want to save the mask only resize the picture2 as same as size of the mask
' or make picture2.autosize = true

Dim S As String
S = Text4.Text
If Picture2.Picture <> 0 Then
    Bmp.SaveMaskToFile Picture2, S
End If
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
On Error GoTo I:
CMD.ShowOpen
sFile = CMD.FileName
If sFile <> "" Then
    If Dir$(sFile) = "" Then GoTo I:
    Picture1.Picture = LoadPicture(sFile)
End If
I:
End Sub

Private Sub Command5_Click()
On Error GoTo I:
CMD.Flags = cdlCCRGBInit
CMD.ShowColor
ColorTrans = CMD.Color
Text1.Locked = False
Text1.Text = CStr(ColorTrans)
Text1.Locked = True
I:
End Sub

Private Sub Command6_Click()
On Error GoTo I:
CMD.Flags = cdlCCRGBInit
CMD.ShowColor
ColorBmp = CMD.Color
Text2.Locked = False
Text2.Text = CStr(ColorBmp)
Text2.Locked = True
I:
End Sub

Private Sub Command7_Click()
On Error Resume Next
CMD.Flags = cdlCCRGBInit
CMD.ShowColor
ColorMask = CMD.Color
Text3.Locked = False
Text3.Text = CStr(ColorMask)
Text3.Locked = True
I:
End Sub

Private Sub Form_Load()
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
ColorTrans = -1
ColorBmp = -1
ColorMask = -1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not Picture3.BackColor = Me.BackColor Then Picture3.BackColor = Me.BackColor
End Sub

Private Sub Picture1_Click()
ColorTrans = EyeDropper(Picture1, PP)
Text1.Locked = False
Text1.Text = CStr(ColorTrans)
Text1.Locked = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
With PP
    .XX = x
    .YY = y
End With
Picture3.BackColor = EyeDropper(Picture1, PP)
End Sub
Private Function EyeDropper(ByRef PIC As PictureBox, ByRef P As POINT) As Long
Dim tmp As Long
tmp = PIC.POINT(PP.XX, PP.YY)
EyeDropper = tmp
End Function

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not Picture3.BackColor = Me.BackColor Then Picture3.BackColor = Me.BackColor
End Sub

