VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Roses"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SaveBtn 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   7200
      Width           =   5775
   End
   Begin VB.CommandButton CreateBtn 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   6120
      Width           =   5775
   End
   Begin VB.PictureBox FlowerPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   5775
      Left            =   240
      ScaleHeight     =   5715
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Pi As Single '3.14
Dim MaxDist As Single 'used for scaling picture
Dim ColDet As Integer 'used to determine which gradient to use
Dim Red As Integer 'initial red value
Dim Green As Integer 'intitial green value
Dim Blue As Integer 'initial blue value
Dim RedVal As Integer 'grdient red value
Dim GreenVal As Integer 'gradient green value
Dim BlueVal As Integer 'gradient blue value
Dim A1 As Single 'constant used generate rose
Dim B1 As Single 'constant used to generate rose

Dim Active As Boolean 'start/stop



Private Sub CreateBtn_Click()
    'starts and stops rose
    Active = Not Active
    
    If Active = True Then
        Me.CreateBtn.Caption = "Stop"
        Me.SaveBtn.Enabled = False
        Me.FlowerPic.Cls
        Get_Seeds
        Get_Base_Colors
        Scale_Pic
        Draw_Rose
    Else
        Me.CreateBtn.Caption = "Start"
        Me.SaveBtn.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Pi = 4 * Atn(1)
    Active = False
    Me.CreateBtn.Caption = "Start"
    Me.SaveBtn.Enabled = False
End Sub


Sub Get_Seeds()
    Dim A As Single
    Dim B As Single
    Dim N As Integer
    
    Randomize
    'Get Random Seeds
    A = Int((32000 * Rnd) + 1)
    N = Int((32000 * Rnd) + 1)
    'I like even numbers
    If N Mod 2 <> 0 Then
        N = N + 1
    End If
    B = A / N
    A1 = (A + 2 * B) / 2
    B1 = (A + B) / B
End Sub

Sub Get_Base_Colors()
    'Get Base Colors For Gradient
    ColDet = Int((8 * Rnd) + 1)
    Red = Int(255 * Rnd)
    Green = Int(255 * Rnd)
    Blue = Int(255 * Rnd)
End Sub

Sub Scale_Pic()
    Dim I As Single
    Dim Phi As Single
    Dim X As Single
    Dim Y As Single
    Dim Dist As Single
    
    'I go through a quick rotaion and
    'determine the max distance from 0
    'I then scale the picture slightly bigger
    MaxDist = 0
    For I = 0 To 360
        Phi = I * Pi / 180
        X = A1 * (Cos(Phi) - Cos(B1 * Phi))
        Y = A1 * (Sin(Phi) - Sin(B1 * Phi))
        
        Dist = Sqr(X ^ 2 + Y ^ 2)
        If Dist > MaxDist Then
            MaxDist = Dist
        End If
        DoEvents
    Next I
    Me.FlowerPic.DrawWidth = 1
    Me.FlowerPic.Scale (-1.01 * MaxDist, MaxDist)-(1.01 * MaxDist, -1.01 * MaxDist)
End Sub

Sub Draw_Rose()
    Dim I As Single 'Angle increment
    Dim J As Single 'gradient increment
    Dim X As Single 'current point
    Dim Y As Single 'current point
    Dim XOld As Single 'old point
    Dim YOld As Single 'old point
    Dim Dist As Single 'distance current pint from 0
    Dim Phi As Single 'angle in radians
    Dim Theta As Single 'angle between old and new point
    Dim GxTemp As Single '(New point) used in gradient, short portion of total line
    Dim GyTemp As Single '(New point) used in gradient, short portion of total line
    Dim GxOld As Single '(Old point) used in gradient, short portion of total line
    Dim GyOld As Single '(Old point) used in gradient, short portion of total line
    Dim DistTemp As Single 'distance between old and new point
    Dim DistInc As Single 'distance to increment gradient
    Dim DistCur As Single 'distance gradient point is from 0
    
    On Error Resume Next
    
    XOld = 0
    YOld = 0
    
    For I = 0 To 360
        
        Phi = I * Pi / 180
        X = A1 * (Cos(Phi) - Cos(B1 * Phi))
        Y = A1 * (Sin(Phi) - Sin(B1 * Phi))
     
       
        'get angle between two point
        Theta = Atn((YOld - Y) / (XOld - X))
        If XOld > X And YOld > Y Then
            Theta = Theta
        ElseIf XOld < X And YOld > Y Then
            Theta = Pi + Theta
        ElseIf XOld < X And YOld < Y Then
            Theta = Pi + Theta
        ElseIf XOld > X And YOld < Y Then
            Theta = 2 * Pi + Theta
        ElseIf XOld > X And YOld = Y Then
            Theta = 0
        ElseIf XOld = X And YOld > Y Then
            Theta = Pi / 2
        ElseIf XOld < X And YOld = Y Then
            Theta = Pi
        ElseIf XOld = X And YOld < Y Then
            Theta = 3 * Pi / 2
        Else
            Theta = 0
        End If
      
        Dist = Sqr(X ^ 2 + Y ^ 2)
        
        GxTemp = X
        GyTemp = Y
        GxOld = X
        GyOld = Y
        DistTemp = Sqr((XOld - X) ^ 2 + (YOld - Y) ^ 2)
        DistInc = DistTemp / 10
        
        'Do gradient from old point to new point
        For J = 1 To 10
            GxTemp = GxTemp + DistInc * Cos(Theta)
            GyTemp = GyTemp + DistInc * Sin(Theta)
            DistCur = Sqr(GxTemp ^ 2 + GyTemp ^ 2)
            Select Case ColDet
                Case 1
                    RedVal = Red * DistCur / MaxDist
                    GreenVal = Green * DistCur / MaxDist
                    BlueVal = Blue * DistCur / MaxDist
                Case 2
                    RedVal = Red * DistCur / MaxDist
                    GreenVal = Green * DistCur / MaxDist
                    BlueVal = Blue - Blue * DistCur / MaxDist
                Case 3
                    RedVal = Red * DistCur / MaxDist
                    GreenVal = Green - Green * DistCur / MaxDist
                    BlueVal = Blue * DistCur / MaxDist
                Case 4
                    RedVal = Red * DistCur / MaxDist
                    GreenVal = Green - Green * DistCur / MaxDist
                    BlueVal = Blue - Blue * DistCur / MaxDist
                Case 5
                    RedVal = Red - Red * DistCur / MaxDist
                    GreenVal = Green * DistCur / MaxDist
                    BlueVal = Blue * DistCur / MaxDist
                Case 6
                    RedVal = Red - Red * DistCur / MaxDist
                    GreenVal = Green * DistCur / MaxDist
                    BlueVal = Blue - Blue * DistCur / MaxDist
                Case 7
                    RedVal = Red - Red * DistCur / MaxDist
                    GreenVal = Green - Green * DistCur / MaxDist
                    BlueVal = Blue * DistCur / MaxDist
                Case 8
                    RedVal = Red - Red * DistCur / MaxDist
                    GreenVal = Green - Green * DistCur / MaxDist
                    BlueVal = Blue - Blue * DistCur / MaxDist
            End Select
            'Draw short portion of total line
            Me.FlowerPic.Line (GxTemp, GyTemp)-(GxOld, GyOld), RGB(RedVal, GreenVal, BlueVal)
            'reset old point to current point
            GxOld = GxTemp
            GyOld = GyTemp
            If Active = False Then
                Exit For
            End If
            DoEvents
        Next J
        'reset old point to current point
        XOld = X
        YOld = Y
        If Active = False Then
            Exit For
        End If
        DoEvents
    Next I
    Me.CreateBtn.Caption = "Start"
    Me.SaveBtn.Enabled = True
    Active = False
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub SaveBtn_Click()
    'saves picture
    On Error Resume Next
    With dlgSave
        .InitDir = "C:\"
        .Filter = "Bitmap (*.bmp)|*.bmp"
        .FileName = "Rose"
        .DialogTitle = "Save Rose"
        .CancelError = True
        .ShowSave
        If Err <> MSComDlg.cdlCancel Then
            Screen.MousePointer = vbHourglass
            SavePicture Me.FlowerPic.Image, .FileName
            Screen.MousePointer = vbDefault
        End If
    End With
    On Error GoTo 0
End Sub
