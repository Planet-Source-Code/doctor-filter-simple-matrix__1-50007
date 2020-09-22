VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form Console 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Matrix"
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   3165
   DrawMode        =   1  'Blackness
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00FF8080&
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   6
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   1800
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
   End
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   1320
      Top             =   0
   End
   Begin VB.PictureBox VDU 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      FillColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Dialog"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000001&
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   3255
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ---------------
'  SIMPLE MATRIX
' ---------------
'
' This code shows you how to create a simple Matrix effect like you see in
' the Matrix films. It's so easy it hurts! I got bored to death writing the
' comments so, tough luck.
'
' It works really simply by creating an array of letter generators that are
' white at the tip and merge into a random dark green colour. These geneators
' run down the screen, leaving the last-darkest green letter behind, kind of
' like a snail trail.
'
' "Simple!"
'
' Please send any comments or ideas to "doctor.filter@virgin.net"
'
'-----------------------------------------------------------------------------
Option Explicit
Dim n                   ' n is used to scroll through an array eg: ONCurX(n)
Dim ONCurX(1 To 100)    ' this is an array for the Horizontal position of the generator of a text strand
Dim ONCurY(1 To 100)    ' this is an array for the Vertical position of the generator of a text strand
Dim OnCol(1 To 100)     ' this is an array for the Shade of Green of a particular strand
Dim OFFCurX(1 To 100)   ' this is an array for Black characters that remove the trails (Horizontal)
Dim OFFCurY(1 To 100)   ' this is an array for Black characters that remove the trails (Vertical)
Dim CharSet$            ' This string will later contain all the characters that will be displayed in the strands
Dim CurCount            ' CurCount Specifies how many strands are drawn at once
Dim OutCount            ' OutCount Specifies how many strands are blackened at once
Dim FontGap             ' This is the distance bewteen letters
Dim TextDivis           ' This is the amount of letters that you can fit into the picture box
Dim DarkGreen
Dim LightGreen
Dim FadeColour
Dim FadeStep
Dim FadeCount
Dim Fades
Dim FadeVal
Dim StrandDiff
Dim TYVAL


Private Sub Form_Load()

Randomize Timer ' Make the Stands start in a different place every time the program starts
Console.Left = SysInfo1.WorkAreaLeft ' Move the form to the left of the screen
Console.Top = SysInfo1.WorkAreaTop   ' Move the form to the top of the screen
Console.Width = SysInfo1.WorkAreaWidth ' Make the form as wide as the screen
Console.Height = SysInfo1.WorkAreaHeight ' Make the form as height as the screen
VDU.Width = SysInfo1.WorkAreaWidth ' Make the picture box the same width as the screen
VDU.Height = SysInfo1.WorkAreaHeight ' Make the picture box the same height as the screen

FontGap = 120 ' This is the distance bewteen letters on the picture box (VDU)
CharSet$ = "!£$%^&*()_+1234567890-=[]#{}~@:';?></.,¬`|\" ' These are the charaters that get generated 'try changing it to your "firstname" or to "10"
CurCount = 30 ' Amount of strands to be drawn on the picture box
FadeCount = 10 ' The length (in characters) between the white tip and the back end of the generator
LightGreen = 255 ' The brightness of the brightest green :lol
DarkGreen = 50 ' The darkness of the darkest green :LMFAO!
FadeStep = (LightGreen - DarkGreen) / FadeCount ' Work out the value for the difference of each faded step of green in the generator
OutCount = 10 'The amount of string deleters removing the strands of code
StrandDiff = 60 ' The maximum varience in colour between each strand of code (not each strand generator)

TextDivis = Int(VDU.Width / FontGap) ' This is the amount of letters that you can fit into the picture box

' Ok, can't be bothered, work the rest out yourself.

For n = 1 To CurCount
    ONCurX(n) = Int(Rnd * TextDivis) * FontGap
    ONCurY(n) = Int(Rnd * VDU.Height)
    OnCol(n) = Int(Rnd * StrandDiff)
   
Next n

For n = 1 To OutCount
    OFFCurX(n) = Int(Rnd * TextDivis) * FontGap
    OFFCurY(n) = Int(Rnd * VDU.Height)
   
Next n

End Sub


Private Sub Timer_Timer()

VDU.FontSize = 6
VDU.Font = "Terminal"
For n = 1 To CurCount
    VDU.CurrentY = ONCurY(n)
    VDU.CurrentX = ONCurX(n)
    VDU.ForeColor = vbBlack
    VDU.Print " "
    VDU.ForeColor = vbWhite
    VDU.CurrentY = ONCurY(n)
    VDU.CurrentX = ONCurX(n)
    VDU.Print Mid$(CharSet$, Int(Rnd * Len(CharSet$) + 1), 1)
    '********************************************************
    FadeStep = (LightGreen - OnCol(n)) / FadeCount
    For Fades = 1 To FadeCount
        VDU.CurrentY = ONCurY(n) - (FontGap * Fades)
        VDU.CurrentX = ONCurX(n)
        VDU.ForeColor = vbBlack
        VDU.Print " "
        VDU.ForeColor = RGB(0, LightGreen - (FadeStep * Fades - 1), 0)
        VDU.CurrentY = ONCurY(n) - (FontGap * Fades)
        VDU.CurrentX = ONCurX(n)
        VDU.Print Mid$(CharSet$, Int(Rnd * Len(CharSet$) + 1), 1)
    Next Fades
    '********************************************************
        TYVAL = ONCurY(n) - 20 * FontGap
        VDU.CurrentY = TYVAL
        VDU.CurrentX = ONCurX(n)
        VDU.ForeColor = vbBlack
        VDU.Print " "
        VDU.ForeColor = RGB(0, OnCol(n), 0)
        VDU.CurrentY = TYVAL
        VDU.CurrentX = ONCurX(n)
        VDU.Print Mid$(CharSet$, Int(Rnd * Len(CharSet$) + 1), 1)
    '********************************************************
    ONCurY(n) = ONCurY(n) + FontGap
    If ONCurY(n) >= VDU.Height + (FontGap * FadeCount) Then
        ONCurY(n) = VDU.Top
        ONCurX(n) = Int(Rnd * TextDivis) * FontGap
        OnCol(n) = Int(Rnd * StrandDiff)
    End If
Next n

For n = 1 To OutCount
    VDU.CurrentY = OFFCurY(n)
    VDU.CurrentX = OFFCurX(n)
    VDU.ForeColor = vbBlack
    VDU.Print " "
    OFFCurY(n) = OFFCurY(n) + FontGap
    If OFFCurY(n) >= VDU.Height Then
        OFFCurY(n) = VDU.Top
        OFFCurX(n) = Int(Rnd * TextDivis) * FontGap
    End If
Next n

End Sub

Private Sub VDU_DblClick()
End
End Sub
