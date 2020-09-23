VERSION 5.00
Object = "*\A..\prjImageList.vbp"
Begin VB.Form frmImlTest 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ImageList Test"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6300
   StartUpPosition =   1  'CenterOwner
   Begin prjImageList.ucTest ucTest1 
      Height          =   555
      Left            =   4590
      TabIndex        =   8
      Top             =   315
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   979
      SmallImageCount =   32
      SmallImages     =   "frmImlTest.frx":0000
      LargeIconSizeX  =   48
      LargeIconSizeY  =   48
      LargeColourDepth=   32
      LargeImageCount =   15
      LargeImages     =   "frmImlTest.frx":80021
      LargeKeys       =   "frmImlTest.frx":BC042
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Ghosted"
      Height          =   195
      Index           =   4
      Left            =   4500
      TabIndex        =   7
      Top             =   4635
      Width           =   960
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Colorize"
      Height          =   195
      Index           =   3
      Left            =   4500
      TabIndex        =   6
      Top             =   4365
      Width           =   960
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Selected"
      Height          =   195
      Index           =   2
      Left            =   4500
      TabIndex        =   4
      Top             =   4095
      Width           =   960
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Disabled"
      Height          =   195
      Index           =   1
      Left            =   4500
      TabIndex        =   3
      Top             =   3825
      Width           =   960
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Normal"
      Height          =   195
      Index           =   0
      Left            =   4500
      TabIndex        =   2
      Top             =   3555
      Value           =   -1  'True
      Width           =   960
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw Small Images"
      Height          =   375
      Index           =   1
      Left            =   4455
      TabIndex        =   1
      Top             =   5490
      Width           =   1635
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw Large Images"
      Height          =   375
      Index           =   0
      Left            =   4455
      TabIndex        =   0
      Top             =   5040
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Draw Style"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4500
      TabIndex        =   5
      Top             =   3285
      Width           =   930
   End
End
Attribute VB_Name = "frmImlTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lMode     As Long
Private m_iCurIdx   As Integer


Private Sub cmdDraw_Click(Index As Integer)

Dim lCt     As Long
Dim lX      As Long
Dim lY      As Long
Dim lState  As Long

    m_iCurIdx = Index
    Select Case Index
    Case 0
        With ucTest1
            lX = 10
            lY = 10
            For lCt = 0 To .LargeImageCount - 1
                If (lY > (Me.ScaleHeight / Screen.TwipsPerPixelY) - .LargeImageY) Then
                    lY = 10
                    lX = lX + .LargeImageX + 4
                End If
                lState = m_lMode
                .Draw Me.hDC, lCt, lX, lY, lState, 35539, True
                lY = lY + .LargeImageY + 4
            Next lCt
        End With
    Case 1
        With ucTest1
            lX = 180
            lY = 10
            For lCt = 0 To .SmallImageCount - 1
                If (lY > (Me.ScaleHeight / Screen.TwipsPerPixelY) - 20) Then
                    lY = 10
                    lX = lX + .SmallImageX + 4
                End If
                lState = m_lMode
                .Draw Me.hDC, lCt, lX, lY, lState, 97035, False
                lY = lY + .SmallImageY + 4
            Next lCt
        End With
    End Select
    
    Me.Refresh

End Sub

Private Sub optMode_Click(Index As Integer)
    m_lMode = Index
    Me.Cls
    cmdDraw_Click m_iCurIdx
End Sub
