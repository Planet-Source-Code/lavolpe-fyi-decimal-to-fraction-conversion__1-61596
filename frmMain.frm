VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCarp 
      Height          =   285
      Left            =   3750
      TabIndex        =   13
      Text            =   "0"
      Top             =   4785
      Width           =   4185
   End
   Begin VB.OptionButton optMeasure 
      Caption         =   "Decimal is in Inches"
      Height          =   210
      Index           =   1
      Left            =   1845
      TabIndex        =   12
      Top             =   4815
      Width           =   1800
   End
   Begin VB.OptionButton optMeasure 
      Caption         =   "Decimal is in Feet"
      Height          =   210
      Index           =   0
      Left            =   90
      TabIndex        =   11
      Top             =   4800
      Value           =   -1  'True
      Width           =   1725
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4410
      TabIndex        =   7
      Text            =   "100"
      Top             =   720
      Width           =   1455
   End
   Begin VB.ComboBox cboDigits 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   2550
      List            =   "frmMain.frx":0034
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   105
      TabIndex        =   2
      Top             =   1050
      Width           =   8910
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fraction"
      Height          =   660
      Left            =   7785
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Text            =   "3.14159265358979"
      Top             =   735
      Width           =   2355
   End
   Begin VB.Label Label1 
      Caption         =   "Carpenter's Cheat to convert decimals to fractions of an Inch (32nds of an Inch is Lowest inch part)"
      Height          =   225
      Index           =   5
      Left            =   1755
      TabIndex        =   10
      Top             =   4485
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   "Bonus Material :)"
      Height          =   225
      Index           =   4
      Left            =   135
      TabIndex        =   9
      Top             =   4485
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Accuracy (max 100)"
      Height          =   225
      Index           =   3
      Left            =   4425
      TabIndex        =   8
      Top             =   480
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Max Denominator Digits"
      Height          =   225
      Index           =   2
      Left            =   2565
      TabIndex        =   6
      Top             =   480
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Decimal to Convert"
      Height          =   225
      Index           =   1
      Left            =   165
      TabIndex        =   4
      Top             =   495
      Width           =   2340
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Optional Parameters"
      Height          =   225
      Index           =   0
      Left            =   2535
      TabIndex        =   3
      Top             =   165
      Width           =   3360
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Original code found at following site:
' http://mathforum.org/library/drmath/view/51910.html

' Very Unique.  I like unique ideas/code & think this one is worthy to pass around

' I found it interesting in the fact that it uses vector logic to determine
' the fraction from a decimal.  It was tweaked by me a bit to allow larger
' fractions, negative values, mixed fractions, and to allow some optional settings
' Note: Most Double declares can be changed to Long if so desired

' The advantages of this code are:
'   1. Probably as accurate as any other routines out there that do it differently
'   2. Probably faster as it does less to get the answer
'       -- notice there is no simplifying fractions like other routines would use
'       -- also notice that there is only 1 loop; others out there may have several
'   3. Options to return more simpler close-estimate fractions
'       -- decimal portion of Sqr(2) as a fraction is closest to:
'           434351420/1048617089; ~100% accuracy
'           but at 99.983% accuracy, routine could return 29/70 << much easier on the eyes
        
' I thought the carpenter's cheat could be useful to convert decimals to U.S. feet & inches
' -- fraction related; so I threw it in for the heck of it


Option Explicit


Private Sub ConvertToFraction(ByVal V As Double, _
            W As Double, N As Double, D As Double, _
            Optional ByVal maxDenomDigits As Byte, _
            Optional ByVal Accuracy As Double = 100#, _
            Optional DisplayText As String)

    Const MaxTerms As Integer = 50          'Limit to prevent infinite loop
    Const MinDivisor As Double = 1E-16      'Limit to prevent divide by zero
    Const MaxError As Double = 1E-50        'How close is enough
    Dim F As Double                         'Fraction being converted
    Dim A As Double 'Current term in continued fraction
    Dim N1 As Double 'Numerator, denominator of last approx
    Dim D1 As Double
    Dim N2 As Double 'Numerator, denominator of previous approx
    Dim D2 As Double
    Dim I As Integer
    Dim T As Double
    Dim maxDenom As Double
    Dim bIsNegative As Boolean
    Dim sDec As String
    
    If maxDenomDigits = 0 Or maxDenomDigits > 17 Then maxDenomDigits = 17
    maxDenom = 10 ^ maxDenomDigits
    If Accuracy > 100 Or Accuracy < 1 Then Accuracy = 100
    Accuracy = Accuracy / 100#
    
    bIsNegative = (V < 0)
    W = Abs(Fix(V))
    'V = Abs(V) - W << subtracting doubles can change the decimal portion by adding more numeral at end
    V = CDbl(Mid$(CStr(Abs(V)), Len(CStr(W)) + 1))
    
    ' check for no decimal or zero
    If V = 0 Then GoTo RtnResult
    
    F = V                       'Initialize fraction being converted
    
    N1 = 1                      'Initialize fractions with 1/0, 0/1
    D1 = 0
    N2 = 0
    D2 = 1

    On Error GoTo RtnResult
    For I = 0 To MaxTerms
        A = Fix(F)              'Get next term
        F = F - A               'Get new divisor
        N = N1 * A + N2         'Calculate new fraction
        D = D1 * A + D2
        N2 = N1                 'Save last two fractions
        D2 = D1
        N1 = N
        D1 = D
                                'Quit if dividing by zero
        If F < MinDivisor Then Exit For

                                'Quit if close enough
        T = N / D               ' A=zero indicates exact match or extremely close
        A = Abs(V - T)          ' Difference btwn actual V and calculated T
        If A < MaxError Then Exit For
                                'Quit if max denominator digits encountered
        If D > maxDenom Then Exit For
                                ' Quit if preferred accuracy accomplished
        If N Then
            If T > V Then T = V / T Else T = T / V
            If T >= Accuracy And Abs(T) < 1 Then Exit For
        End If
        F = 1# / F               'Take reciprocal
    Next I

RtnResult:
If Err Or D > maxDenom Then
    ' in above case, use the previous best N & D
    If D2 = 0 Then
        N = N1
        D = D1
    Else
        D = D2
        N = N2
    End If
End If

' correct for negative values
If bIsNegative Then
    If W Then W = -W Else N = -N
End If

' Set this up anyway you want
DisplayText = N & " / " & D
If W Then DisplayText = W & " & " & DisplayText

End Sub


Private Sub Command1_Click()
Dim N As Double, D As Double, W As Double
Dim Target As Double, Accuracy As Double, tResult As Double, T As Double
Dim I As Integer, sDisplay As String

List1.Clear
Target = CDbl(Val(Text1))
Text1 = Target

ConvertToFraction Target, W, N, D, Val(cboDigits.Text) * 0, Val(Text2), sDisplay
If D = 0 Then
    List1.AddItem "No decimal to convert"
    Exit Sub
End If

    ' calculate accuacy. Due to rounding & 16 digit limit for Doubles,
    ' 100% may be reported even though the accuracy may truly be slightly less.
    ' For example. The routine returns the folowing result for PI
    '   PI = 3.14159265358979
    '   Result: 3 & 35580937/251290841
    '   True accuracy via Calculator is: 99.999999999999994697694442392003% not 100%
    
    tResult = N / D
    T = CDec(Mid$(CStr(Target), Len(CStr(Fix(Target))) + 1))
    If Abs(tResult) < Abs(T) Then
        Accuracy = Abs(CDec(tResult / T) * 100#)
    Else
        Accuracy = Abs(CDec(T / tResult) * 100#)
    End If
    If W < 0 Then tResult = (Abs(W) + tResult) * -1 Else tResult = W + tResult
    ' add the 1st return from the function
    List1.AddItem sDisplay & "  = " & tResult & "  Accuracy: " & Left$(Accuracy, 20) & "%"
    


' ////////////////////////////////////////////////////////////
' This portion is simply to return other possible results that
' have smaller denominator values

I = Len(CStr(D))    ' start w/this number of denominator digits
Do Until I < 1

    I = I - 1       ' subtract one & call function
    ConvertToFraction Target, W, N, D, I + 0, , sDisplay
    
    If Len(CStr(D)) < I + 1 Then
        ' calculate accuacy. Due to rounding 100% may be reported even though
        ' the accuracy may truly be something like .999999999999999999999999999
        tResult = N / D
        T = CDec(Mid$(CStr(Target), Len(CStr(Fix(Target))) + 1))
        If Abs(tResult) < Abs(T) Then
            Accuracy = Abs(CDec(tResult / T)) * 100#
        Else
            Accuracy = Abs(CDec(T / tResult)) * 100#
        End If
        If W < 0 Then tResult = (Abs(W) + tResult) * -1 Else tResult = W + tResult
        ' add the 1st return from the function
        List1.AddItem sDisplay & "  = " & tResult & "  Accuracy: " & Left$(Accuracy, 20) & "%"
        I = Len(CStr(D))    ' update denominator digits
    End If

Loop

DoCarpenterCheat

End Sub

Private Sub DoCarpenterCheat()

' ////////////////////////////////////////////////////////////
' Carpenter's Cheat...
Dim Ft As Long
Dim Inch As Integer, I As Integer
Dim Target As Double
Dim tmp As Double
Dim sDisplay As String, sInchText As String

Target = Abs(CDbl(Val(Text1)))

If optMeasure(0) = True Then            ' feet vs inches
    Ft = Abs(Fix(Target))               ' whole number is feet
    tmp = (Abs(Target) - Ft) * 12       ' calc percentage of foot for decimal
    Inch = Fix(tmp)                     ' this will be inches
    I = CInt(32 * (tmp - Inch))         ' calc percentage of inches left
Else
    Inch = Fix(Target)                  ' whole number is inches
    I = CInt(32 * (Target - Inch))      ' calc percentage of inches left
End If

If I = 32 Then                          ' if=32 then we got a whole inch
    Inch = Inch + 1
    If Inch = 12 Then                   ' if nr of inches=12 then we got a whole foot
        Ft = Ft + 1
        Inch = 0
    End If
ElseIf I > 0 Then                       ' break down the 1/32 to larger fraction
    tmp = 32                            ' if possible
    Do Until I Mod 2 > 0
        I = I / 2
        tmp = tmp / 2
    Loop
    sDisplay = I & " / " & tmp          ' start building the display text
End If

' The result is done & basically looks like this....
'     ft : inch : I/tmp     where tmp is either 32,16,8,4,2

If Inch Then sInchText = Inch & " & "   ' include/exclude the "Inch" portion if Inch=0

' finalize the display text & display it
If Ft > 0 Or optMeasure(0) = True Then
    sDisplay = Target & " =~ " & Ft & "' and " & sInchText & sDisplay & " inches"
Else
    sDisplay = Target & " =~ " & sInchText & sDisplay & " inches"
End If
txtCarp = sDisplay

End Sub

Private Sub Form_Load()
cboDigits.ListIndex = cboDigits.ListCount - 1
optMeasure(1) = True
End Sub

Private Sub optMeasure_Click(Index As Integer)
DoCarpenterCheat
End Sub
