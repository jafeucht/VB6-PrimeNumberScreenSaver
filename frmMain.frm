VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   4785
   ClientLeft      =   1695
   ClientTop       =   2025
   ClientWidth     =   4500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   2
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmTimer 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Array of found prime numbers
Dim Primes() As Long

Private Type PointAPI
    X As Single
    Y As Single
End Type

' All the Win32 video declares and whatnot.
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Dim Mov As PointAPI
Dim PrintTxt As Boolean

Private Declare Function VerifyScreenSavePwd Lib "password.cpl" (ByVal hwnd&) As Boolean

Function PrimeCnt() As Long
    ' Return the number of entries in the Prime array
    On Error Resume Next
    PrimeCnt = UBound(Primes)
End Function

Sub AddPrime(NewPrime As Long)
    ' Add a new value to the array of primes
    ReDim Preserve Primes(1 To PrimeCnt + 1)
    Primes(PrimeCnt) = NewPrime
End Sub

Function FindPrime() As Long
Dim CurNum As Long, FoundPrime As Boolean, CntInt As Integer, i As Integer
    ' Start counting up from the last found prime number
    If PrimeCnt > 0 Then CurNum = Primes(PrimeCnt)
    ' Count by 1
    CntInt = 1
    ' Start with 2
    If CurNum = 0 Then CurNum = 1
    ' 2 is the only even prime number, so count by 2 after it
    If CurNum > 2 Then CntInt = 2
    ' Loop until multiple is found
    Do
        ' Assume current number is prime
        FoundPrime = True
        ' Increment current number
        CurNum = CurNum + CntInt
        ' Loop through other found prime numbers to see if there is
        ' any multiple in the current number
        For i = 1 To PrimeCnt
            ' If the current number is divisible by a prime number,
            ' then it is not a prime number.
            If IsLng(CurNum / Primes(i)) Then
                FoundPrime = False
                Exit For
            End If
            ' If the current prime number is greater than the square
            ' root of the current number, then curnum is a prime number.
            If Sqr(CurNum) < Primes(i) Then FoundPrime = True
        Next i
    Loop Until FoundPrime
    ' Add the prime number to the prime database
    AddPrime CurNum
    ' Return the found prime number
    FindPrime = CurNum
End Function

Function IsLng(Number As Double) As Boolean
    ' Return false if it is a decimal
    IsLng = (Number = CLng(Number))
End Function

Private Sub Animate()
Dim NewPrime As String, TxTimer As Double, TickCnt As Long
    ' This proceedure is the core of all graphic modification to the form
    Do
        DoEvents
        TickCnt = TickCnt + 1
        TxTimer = TxTimer + 1
        ' After so many loops, allow a prime number to be
        ' printed to the form
        If TxTimer >= TxtInt Then
            PrintTxt = True
            TxTimer = 0
        End If
        ' If it is time to print a prime number, then...
        If PrintTxt Then
            PrintTxt = False
            ' Find the next prime number
            NewPrime = FindPrime
            ' Create a random color
            ForeColor = RGB(Int(Rnd * 155) + 100, Int(Rnd * 155) + 100, Int(Rnd * 155) + 100)
            ' Create a random print position
            CurrentX = Int(Rnd * (ScaleWidth - TextWidth(NewPrime) - MovLen * 2) + MovLen)
            CurrentY = Int(Rnd * (ScaleHeight - TextHeight(NewPrime) - MovLen * 2) + MovLen)
            ' Print the new prime number to the screen
            Print NewPrime
        End If
        ' After 2000 clicks, ...
        If TickCnt = 2000 Then
            TickCnt = 0
            ' Change the motion of the prime numbers
            If Mov.X = Mov.Y Then
                Mov.X = -Mov.X
            Else
                Mov.Y = -Mov.Y
            End If
        End If
        ' Move the prime numbers
        MovScreen
    Loop
End Sub

Private Sub MovScreen()
    ' Move the entire screen over a little bit
    If MovLen = 0 Then Exit Sub
    Picture = Image
    Line (1, 1)-(ScaleWidth - 1, ScaleHeight - 1), 0, B
    PaintPicture Picture, Mov.X, Mov.Y, , , , , , , vbSrcErase Or vbSrcCopy
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FormActivity
End Sub

Private Sub Form_Load()
Dim i As Integer
    If Not RunMode = Preview Then
        
        If UsePassword Then
            ' Disable other screens
            Call SystemParametersInfo(97, 1, 0, 0)
            ' Set window as topmost
            SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
        End If
        
        'Hide the mouse
        ShowCurs False
    End If
    
    ' Create random direction
    Mov.X = (Int(Rnd * 2) * 2 - 1)
    Mov.Y = (Int(Rnd * 2) * 2 - 1)
    
    ' Set up the form font
    FontName = FtName
    FontBold = FtBold
    FontItalic = FtItalic
    FontUnderline = FtUnderline
    FontSize = FtSize
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormActivity
End Sub

Sub FormActivity()
' Proceedure run whenever mouse/keyboard activity happens
Static Count As Integer
        
    Count = Count + 1 ' Give enough time for program to run
    
    If Count > 5 Then
        If RunMode = ScreenSaver Then
            Unload Me
        End If
    End If
End Sub

Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormActivity
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'If Windows is shut down close this application too
    If UnloadMode = vbAppWindows Then Exit Sub
    
    ShowCurs True
    'if a password is beeing used ask for it and check its validity
    If RunMode = ScreenSaver And UsePassword Then
        ' Verify password
        If (VerifyScreenSavePwd(hwnd)) = False Then
            ' Wrong password
            Cancel = True
        End If
    End If
    
    If Not Cancel Then End
    
    ShowCurs False

End Sub

Private Sub tmTimer_Timer()

    
    
    If RunMode = Preview Then
        ' Adjust for smaller screen
        Font.Size = 6
    Else
        ' Adjust for larger screen
        Mov.X = Mov.X * MovLen
        Mov.Y = Mov.Y * MovLen
    End If
    
    ' Start animation
    Animate
    
    tmTimer.Enabled = False
End Sub
