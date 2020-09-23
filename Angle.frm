VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "aNGLES BY yOVAS(c) 2000"
   ClientHeight    =   5865
   ClientLeft      =   2940
   ClientTop       =   1440
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   391
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   Begin VB.PictureBox za 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox az 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   4920
      Width           =   615
   End
   Begin MSComDlg.CommonDialog Cd 
      Left            =   120
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "320"
      Top             =   45
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "gO"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "INSERT ANGLE"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   75
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please if you use my code give me credit
'and use it in projects no-commercials
'Yovas(c) 2000


Option Explicit

Dim Xinc As Single
Dim Yinc As Single
Dim AnG As Integer
Const PhI = 3.14159265358979

Sub Slide(Img As PictureBox)
 
    Do
    
    Call Rebote(az, za)
                       
    If Img.Left >= Form1.ScaleWidth - Img.Width Then
            
                        
            If (AnG <= 360 And AnG >= 270) Then
                AnG = 180 + (360 - Abs(AnG))
            ElseIf AnG <= 90 And AnG >= 0 Then
                AnG = Abs(180 - AnG)
            End If
            
        Call gO
    End If
    
    If Img.Top >= Form1.ScaleHeight - Img.Height Then
            If AnG <= 360 And AnG >= 270 Then
                AnG = 360 - Abs(AnG)
            ElseIf AnG < 270 And AnG >= 180 Then
                AnG = 360 - Abs(AnG)
            End If
            
        Call gO
    End If
    
    If Img.Left <= 0 Then
            If AnG <= 180 And AnG >= 90 Then
                 AnG = Abs(180 - Abs(360 - AnG))
            ElseIf AnG <= 270 And AnG > 180 Then
                AnG = 180 + Abs(360 - AnG)
            End If
            
        Call gO
    End If
    
    If Img.Top <= 0 Then
            If AnG <= 90 And AnG >= 0 Then
                 AnG = 360 - Abs(AnG)
            ElseIf AnG <= 180 And AnG > 90 Then
                AnG = 360 - Abs(AnG)
            End If
            
        Call gO
    End If
           
    DoEvents
    
    Img.Left = Img.Left + Xinc
    Img.Top = Img.Top - Yinc
    
    Loop

    
    
End Sub

Sub gO()
        
    Xinc = Cos(AnG * PhI / 180) / 5
    Yinc = Sin(AnG * PhI / 180) / 5
            
       If AnG > 360 Then AnG = Abs(AnG - 360)
       If AnG < 0 Then AnG = 360 + AnG
           
End Sub

Private Sub Command1_Click()
    AnG = Text1.Text
    Command1.Visible = False
    Text1.Visible = False
    Label1.Visible = False
    Call gO
    Call Slide(az)
End Sub


Sub Rebote(Img1 As PictureBox, Img2 As PictureBox)

If (Img1.Left >= Img2.Left - Img1.Width) And (Img1.Left <= Img2.Left + Img2.Width) Then
        If (Img1.Top >= Img2.Top - Img1.Height) And (Img1.Top <= Img2.Top + Img2.Height) Then
            Img2.Visible = False
            Abracadabra Img2
        End If
                
    End If

End Sub

Sub Abracadabra(Img As PictureBox)
Dim laY As Single
Dim laX As Single

la_Suerte:

    Randomize
    laX = (Rnd * Form1.ScaleWidth) * 5 + 10
    laY = (Rnd * Form1.ScaleHeight) * Rnd * 2
    
    If (laX < Img.Width Or laX > Form1.ScaleWidth - Img.Width) Or (laY < Img.Height Or laY > Form1.ScaleHeight - Img.Height) Then GoTo la_Suerte
    
    Img.Left = laX
    Img.Top = laY
    
    Img.Visible = True
    
End Sub

Private Sub Form_Load()

Loading:

On Error GoTo Shit

MsgBox "Angles by Yovas(c) 2000", , "Angles"

MsgBox "Please open an image (not too big please!)", , "Angles"

Cd.Filter = "Archivos de imagen |*.jpg;*.bmp;*.gif;*.ico;*.jpeg|"
Cd.FileName = ""
Cd.ShowOpen
If Not Trim(Cd.FileName) <> "" Then GoTo Loading
az.Picture = LoadPicture(Cd.FileName)

MsgBox "Please open another image (not too big please!)", , "Angles"
Cd.FileName = ""
Cd.ShowOpen
If Not Trim(Cd.FileName) <> "" Then GoTo Loading
za.Picture = LoadPicture(Cd.FileName)

Form1.WindowState = 2
Exit Sub

Shit:
End

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
