VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "BoxShot Demo by Ken Foster"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HS1 
      Height          =   225
      Left            =   7335
      Max             =   8
      Min             =   2
      TabIndex        =   27
      Top             =   7110
      Value           =   4
      Width           =   2370
   End
   Begin VB.CheckBox chkOptions2 
      Caption         =   "Shadow"
      Height          =   195
      Index           =   2
      Left            =   4740
      TabIndex        =   26
      Top             =   6105
      Width           =   915
   End
   Begin VB.CommandButton cmdSaveJpeg 
      Caption         =   "Save as Jpeg"
      Height          =   435
      Left            =   8460
      TabIndex        =   25
      Top             =   6330
      Width           =   1395
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7635
      TabIndex        =   23
      Top             =   5880
      Width           =   1920
   End
   Begin VB.CheckBox chkOptions2 
      Caption         =   "Left View"
      Height          =   330
      Index           =   1
      Left            =   4740
      TabIndex        =   10
      Top             =   5760
      Width           =   975
   End
   Begin VB.CheckBox chkOptions2 
      Caption         =   "Smooth"
      Height          =   330
      Index           =   0
      Left            =   4740
      TabIndex        =   9
      Top             =   6345
      Width           =   870
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save as Bitmap"
      Height          =   435
      Left            =   6780
      TabIndex        =   8
      Top             =   6330
      Width           =   1350
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5340
      Left            =   4740
      ScaleHeight     =   354
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   338
      TabIndex        =   4
      Top             =   330
      Width           =   5100
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Pictures"
      Height          =   8070
      Left            =   150
      TabIndex        =   0
      Top             =   255
      Width           =   4470
      Begin VB.CommandButton cmdSetFont 
         Caption         =   "Show Font Prop"
         Enabled         =   0   'False
         Height          =   360
         Left            =   1680
         TabIndex        =   20
         Top             =   6795
         Width           =   1470
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Text to Front Panel"
         Height          =   285
         Left            =   105
         TabIndex        =   17
         Top             =   6060
         Width           =   2100
      End
      Begin VB.CommandButton cmdAddText 
         Caption         =   "Set Text"
         Enabled         =   0   'False
         Height          =   360
         Left            =   75
         TabIndex        =   12
         Top             =   6795
         Width           =   1290
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         TabIndex        =   11
         Top             =   6435
         Width           =   2430
      End
      Begin VB.CommandButton cmdLoadPic 
         Caption         =   "Front Pic"
         Height          =   465
         Index           =   0
         Left            =   1650
         TabIndex        =   7
         Top             =   5430
         Width           =   1170
      End
      Begin VB.CommandButton cmdLoadPic 
         Caption         =   "Side Pic"
         Height          =   465
         Index           =   2
         Left            =   105
         TabIndex        =   6
         Top             =   5430
         Width           =   1170
      End
      Begin VB.CommandButton cmdLoadPic 
         Caption         =   "Top Pic"
         Height          =   465
         Index           =   1
         Left            =   3105
         TabIndex        =   5
         Top             =   5415
         Width           =   1170
      End
      Begin VB.PictureBox picFront 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4350
         Left            =   1140
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   288
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   202
         TabIndex        =   3
         Top             =   975
         Width           =   3060
         Begin VB.Shape Shape1 
            BorderStyle     =   3  'Dot
            Height          =   270
            Left            =   45
            Top             =   60
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   285
            TabIndex        =   13
            Top             =   2595
            Visible         =   0   'False
            Width           =   45
         End
      End
      Begin VB.PictureBox picSide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4350
         Left            =   90
         Picture         =   "Form1.frx":3FB3
         ScaleHeight     =   288
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   2
         Top             =   975
         Width           =   945
      End
      Begin VB.PictureBox picTop 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1140
         Picture         =   "Form1.frx":6339
         ScaleHeight     =   39
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   202
         TabIndex        =   1
         Top             =   285
         Width           =   3060
      End
      Begin VB.Label Label19 
         Caption         =   "Italics:"
         Height          =   195
         Left            =   3450
         TabIndex        =   37
         Top             =   6930
         Width           =   465
      End
      Begin VB.Label Label18 
         Caption         =   "False"
         Height          =   195
         Left            =   3945
         TabIndex        =   36
         Top             =   6945
         Width           =   420
      End
      Begin VB.Label Label17 
         Caption         =   "False"
         Height          =   195
         Left            =   3945
         TabIndex        =   35
         Top             =   7140
         Width           =   405
      End
      Begin VB.Label Label16 
         Caption         =   "Bold:"
         Height          =   210
         Left            =   3540
         TabIndex        =   34
         Top             =   7140
         Width           =   420
      End
      Begin VB.Label Label15 
         Caption         =   "False"
         Height          =   180
         Left            =   3960
         TabIndex        =   33
         Top             =   7575
         Width           =   465
      End
      Begin VB.Label Label14 
         Caption         =   "False"
         Height          =   180
         Left            =   3945
         TabIndex        =   32
         Top             =   7350
         Width           =   465
      End
      Begin VB.Label Label13 
         Caption         =   "StrikeOut:"
         Height          =   195
         Left            =   3195
         TabIndex        =   31
         Top             =   7575
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "UnderLine:"
         Height          =   210
         Left            =   3120
         TabIndex        =   30
         Top             =   7350
         Width           =   825
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4020
         TabIndex        =   22
         Top             =   6495
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "Current Font"
         Height          =   210
         Left            =   45
         TabIndex        =   21
         Top             =   7275
         Width           =   915
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tahoma"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   975
         TabIndex        =   19
         Top             =   7245
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Font Color"
         Height          =   255
         Left            =   3225
         TabIndex        =   18
         Top             =   6720
         Width           =   810
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4035
         TabIndex        =   16
         Top             =   6750
         Width           =   300
      End
      Begin VB.Label Label1 
         Caption         =   "Drag label to desired position and press Set Text button"
         Height          =   255
         Left            =   45
         TabIndex        =   15
         Top             =   7785
         Width           =   4125
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   4395
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Label Label5 
         Caption         =   "Font Size"
         Height          =   240
         Left            =   3285
         TabIndex        =   14
         Top             =   6495
         Width           =   720
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Preview Size----Reduced by"
      Height          =   195
      Left            =   7350
      TabIndex        =   29
      Top             =   6870
      Width           =   1965
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   9390
      TabIndex        =   28
      Top             =   6855
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1410
      Left            =   4725
      Stretch         =   -1  'True
      Top             =   6810
      Width           =   1305
   End
   Begin VB.Label Label9 
      Caption         =   "Save as: (name only)"
      Height          =   420
      Left            =   6735
      TabIndex        =   24
      Top             =   5835
      Width           =   825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   
'used to move label
Private OldX As Integer
Private OldY As Integer
   
   'used to smooth
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
   
Private Type ColorRGBType
   Red As Integer
   Blue As Integer
   Green As Integer
End Type

Private Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, _
lpPoint As POINTAPI, _
ByVal hdcSrc As Long, _
ByVal nXSrc As Long, _
ByVal nYSrc As Long, _
ByVal nWidth As Long, _
ByVal nHeight As Long, _
ByVal hbmMask As Long, _
ByVal xMask As Long, _
ByVal yMask As Long) As Long

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Dim LastDir As String
Dim Ft(2) As POINTAPI  ' This holds our points for plotting
Dim Sd(2) As POINTAPI
Dim Tp(2) As POINTAPI
Dim ret As Long         ' This is the return value of the PlgBlt function

Private Sub Form_Load()
   Call DrawPic
   SelectFont.FontName = "Tahoma"
   SelectFont.Fontsize = 10
End Sub

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      Label4.top = 150
      Label4.Left = 50
      Shape1.top = 150
      Shape1.Left = 50
      Shape1.Visible = True
      Label4.Visible = True
      Text1.Enabled = True
      cmdAddText.Enabled = True
      cmdSetFont.Enabled = True
      Text1.SetFocus
   Else
      Label4.Visible = False
      Shape1.Visible = False
      Text1.Enabled = False
      cmdAddText.Enabled = False
      cmdSetFont.Enabled = False
   End If
End Sub

Private Sub chkOptions2_Click(Index As Integer)
   DrawPic
End Sub

Private Sub cmdAddText_Click()
   If Text1.Text = "" Then
      Label4.Visible = False
      Shape1.Visible = False
      Check1.Value = 0
      Exit Sub
   End If
   With SelectFont
   picFront.Fontsize = .Fontsize
   picFront.ForeColor = Label2.BackColor
   picFront.Font = .FontName
   picFront.FontBold = .FontBold
   picFront.FontUnderline = .FontUnderline
   picFront.FontStrikethru = .FontStrikethru
   picFront.CurrentX = Label4.Left
   picFront.CurrentY = Label4.top
   Label4.Visible = False
   Shape1.Visible = False
   End With
   
   picFront.Print Text1.Text
   Call DrawPic
   Text1.Text = ""
   Check1.Value = 0
End Sub

Private Sub cmdSave_Click()    'save as bitmap
   picMain.picture = picMain.Image
   If Text2.Text = "" Then
      MsgBox "Please enter a name for the file."
      Exit Sub
   Else
      SavePicture picMain, App.Path & "\" & Text2.Text & ".bmp"
      MsgBox "Picture Saved..." & App.Path & "\" & Text2.Text & ".bmp"
   End If
End Sub

Private Sub cmdSaveJpeg_Click()
   picMain.picture = picMain.Image
   If Text2.Text = "" Then
      MsgBox "Please enter a name for file"
   Else
      SaveJPG picMain.picture, App.Path & "\" & Text2.Text & ".jpeg", 80
      MsgBox "Picture Saved  ..." & App.Path & "\" & Text2.Text & ".jpeg"
   End If
End Sub

Private Sub cmdSetFont_Click()

   With SelectFont
     If ShowFont = False Then Exit Sub     'Cancel was clicked
      Label2.BackColor = .FontColor
      Label6.Caption = .FontName
      Label8.Caption = .Fontsize
      Label14.Caption = .FontUnderline
      Label15.Caption = .FontStrikethru
      Label17.Caption = .FontBold
      Label18.Caption = .FontItalic
   End With
   Text1.SetFocus
End Sub

Private Sub cmdLoadPic_Click(Index As Integer)
   On Error Resume Next
   Dim SFile As String
   
   SFile = ShowOpen(Me)
   If SFile = "" Then Exit Sub
   Select Case Index
         Case 0: picFront.picture = LoadPicture(SFile)   ' Load the picture
         Size_AspectPicture picFront, SFile, 3
         Case 1: picTop.picture = LoadPicture(SFile)  ' Load the picture
         Size_AspectPicture picTop, SFile, 2.1
         Case 2: picSide.picture = LoadPicture(SFile)  ' Load the picture
         Size_AspectPicture picSide, SFile, 3
   End Select
 
   Call DrawPic    ' Draw new picture on the screen
End Sub

Public Sub DrawPic()
   Dim xp As Integer
   Dim psh As Integer
   
   picMain.picture = LoadPicture()
   
   If chkOptions2(1).Value = Unchecked Then    'Right View
   'Front                                         Side                                Top
   Ft(0).X = 107:                       Sd(0).X = 80:               Tp(0).X = 80
   Ft(0).Y = 70:                         Sd(0).Y = 50:                Tp(0).Y = 50
   
   Ft(1).X = 285:                      Sd(1).X = 107:              Tp(1).X = 258
   Ft(1).Y = 40:                        Sd(1).Y = 70:                 Tp(1).Y = 20
   
   Ft(2).X = 107:                     Sd(2).X = 80:                 Tp(2).X = 107
   Ft(2).Y = 340:                     Sd(2).Y = 320:                Tp(2).Y = 70
   
Else                         'Left View
   
   'Front                                      Side                                  Top
   Ft(0).X = 67:                      Sd(0).X = 245:                Tp(0).X = 100
   Ft(0).Y = 40:                      Sd(0).Y = 70:                   Tp(0).Y = 20
   
   Ft(1).X = 245:                    Sd(1).X = 278:                Tp(1).X = 278
   Ft(1).Y = 70:                      Sd(1).Y = 49:                   Tp(1).Y = 50
   
   Ft(2).X = 67:                      Sd(2).X = 245:                Tp(2).X = 67
   Ft(2).Y = 310:                    Sd(2).Y = 340:                 Tp(2).Y = 40
   
End If

picMain.Cls

ret = PlgBlt(picMain.hDC, Ft(0), picFront.hDC, 0, 0, picFront.ScaleWidth, picFront.ScaleHeight, 0, 0, 0)
ret = PlgBlt(picMain.hDC, Sd(0), picSide.hDC, 0, 0, picSide.ScaleWidth, picSide.ScaleHeight, 0, 0, 0)
ret = PlgBlt(picMain.hDC, Tp(0), picTop.hDC, 0, 0, picTop.ScaleWidth, picTop.ScaleHeight, 0, 0, 0)
picMain.Refresh

'add shadow
picMain.ForeColor = &HF0F0F0      '&H808080

If chkOptions2(1).Value = Unchecked And chkOptions2(2).Value = Checked Then        'Right view
   For psh = 1 To 40
      picMain.Line (40, 288)-(80, 320 - psh)
   Next psh
End If

If chkOptions2(1).Value = Checked And chkOptions2(2).Value = Checked Then
   For psh = 1 To 38                                                             'left view
      picMain.Line (278, 280 + psh)-(320, 290)
   Next psh
End If

If chkOptions2(0).Value = Checked Then Smooth
Image1.picture = picMain.Image
End Sub

Private Sub Smooth()
   Dim X As Integer
   Dim Y As Integer
   Dim Pixel As Long
   Dim Pixel2 As Long
   Dim Col As ColorRGBType
   Dim pp As Integer
   picMain.picture = picMain.Image
   With picMain
      For X = 0 To .ScaleWidth - 1
         For Y = 6 To .ScaleHeight - 1
            
            Pixel = GetPixel(.hDC, X, Y)
            
            If X < .ScaleWidth - 3 Then
               Pixel2 = GetPixel(.hDC, X + 2, Y)
            End If
            
            Col.Red = (RgbColor(Pixel).Red + RgbColor(Pixel2).Red) / 2
            Col.Green = (RgbColor(Pixel).Green + RgbColor(Pixel2).Green) / 2
            Col.Blue = (RgbColor(Pixel).Blue + RgbColor(Pixel2).Blue) / 2
            
            SetPixelV .hDC, X + 1, Y, RGB(Col.Red, Col.Green, Col.Blue)
            
         Next Y
         .Refresh
      Next X
   End With
End Sub

Private Sub Size_AspectPicture(picture As PictureBox, fname As String, Optional MaxWH As Single = 2)
   Dim Photo As StdPicture
   Dim MaxHeight As Single
   Dim MaxWidth As Single
   Dim sFactor As Single
   
   picture.ScaleMode = 1
   picture.AutoRedraw = True
   Set Photo = LoadPicture(fname)
   
   MaxHeight = 1440 * MaxWH        '1.00 inch
   MaxWidth = 1440 * MaxWH        '1.00 inch
   
   'get aspect ratio
   If Photo.Height > Photo.Width Then
      'if height is more than width
      'get Scale factor based on height
      sFactor = Photo.Height / MaxHeight
   ElseIf Photo.Height < Photo.Width Then
      'if height is less than width
      'get Scale factor based on height
      sFactor = Photo.Width / MaxWidth
   ElseIf Photo.Height = Photo.Width Then
      picture.Height = Photo.Height
      picture.Width = Photo.Height
   End If
   
   picture.Width = (Photo.Width / sFactor)
   picture.Height = (Photo.Height / sFactor)
   
   picture.PaintPicture Photo, 0, 0, picture.Width, picture.Height
   picture.ScaleMode = 3
End Sub

Private Function RgbColor(Color As Long) As ColorRGBType
   RgbColor.Red = (Int(Color And 255)) And 255
   RgbColor.Green = (Int(Color / 256)) And 255
   RgbColor.Blue = (Int(Color / 65536)) And 255
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub HS1_Change()
   Image1.Height = picMain.Height / HS1.Value
   Image1.Width = picMain.Width / HS1.Value
   Label10.Caption = HS1.Value
End Sub

Private Sub HS1_Scroll()
   HS1_Change
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   OldX = X
   OldY = Y
   Shape1.top = Label4.top
   Shape1.Left = Label4.Left
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   picFront.ScaleMode = 1
   If Button = 1 Then
      Label4.Left = Label4.Left + (X - OldX)
      Label4.top = Label4.top + (Y - OldY)
      Shape1.Left = Label4.Left
      Shape1.top = Label4.top
   End If
   picFront.ScaleMode = 3
End Sub

Private Sub Text1_Change()
   With SelectFont
   Label4.Fontsize = .Fontsize
   Label4.ForeColor = .FontColor
   Label4.Font = .FontName
   Label4.FontBold = .FontBold
   Label4.FontUnderline = .FontUnderline
   Label4.FontStrikethru = .FontStrikethru
   Label4.Caption = Text1.Text
   Shape1.Width = Label4.Width
   Shape1.Height = Label4.Height
   End With
End Sub

