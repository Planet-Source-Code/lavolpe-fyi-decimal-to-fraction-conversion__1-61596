VERSION 5.00
Begin VB.Form frmAPIparser 
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox mockListBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Left            =   240
      ScaleHeight     =   3180
      ScaleWidth      =   5745
      TabIndex        =   5
      Top             =   1545
      Width           =   5805
      Begin VB.VScrollBar lbScroll 
         Height          =   3165
         Left            =   5460
         Max             =   1000
         TabIndex        =   6
         Top             =   15
         Width           =   270
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "!"
      Height          =   270
      Left            =   5745
      TabIndex        =   3
      Top             =   570
      Width           =   330
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmAPIparser.frx":0000
      Left            =   240
      List            =   "frmAPIparser.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1065
      Width           =   5760
   End
   Begin VB.TextBox Text2 
      Height          =   1830
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmAPIparser.frx":0042
      Top             =   4830
      Width           =   5880
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "C:\Program Files\ApiViewer 2004\Win32api.apv"
      Top             =   570
      Width           =   5430
   End
   Begin VB.Label Label1 
      Caption         =   "Change apv file to Win16api, Win32api or WinCEapi.apv then hit the ! button"
      Height          =   270
      Left            =   210
      TabIndex        =   4
      Top             =   270
      Width           =   5580
   End
End
Attribute VB_Name = "frmAPIparser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Skeleton of api viewer parser.

' Compatible with v2 & v3 files for Win16, Win32 & WinCE versions

' I would think that building cached index array (A-Z) to identify which
' file bytes the items start on for that specific letter would speed up things
' much. The api viewer is not indexed; therefore, you will need to create your
' own.

' Core indexes are created in this sample to cache where specific sections
' of the api viewer begin.

' Note that I open & close the file for each action. You could cache the filenumber
' and simply close when your app closes.

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Const DT_CALCRECT As Long = &H400
Private Const DT_VCENTER As Long = &H4
Private Const DT_SINGLELINE As Long = &H20
Private mockTopIndex As Long
Private mockListIndex As Long
Private mockListCount As Long
Private mockListItems() As String
Private mockScrollRatio As Single
Private hBrushFill(0 To 1) As Long
Private bScrolling As Boolean

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Enum apiSections
    apDeclarations = 0
    apTypes = 1
    apConstants = 2
    apEnumerations = 3
End Enum
Private apiTotals(0 To 3) As Long   ' number in each section
Private apiOffsets(0 To 7) As Long  ' byteoffset in the apv file
'0=declare names
'1=declare definitions
'2=type names
'3=type members
'4=constants
'5=constant values
'6=enumerations
'7=enumeration members

Private Sub Combo1_Click()
'GetAPIlisting Combo1.ListIndex
Call mockListBox_Resize
End Sub

Private Sub Command1_Click()
'List1.Clear
Call mockListBox_Resize
InitializeOffsets
Call Combo1_Click
'If List1.ListCount Then List1.ListIndex = 0
End Sub

Private Sub Form_Load()
If InitializeOffsets() Then
    Combo1.ListIndex = 0
    'If List1.ListCount Then List1.ListIndex = 0
    'Call mockListBox_Resize
End If
End Sub

'Private Sub List1_Click()
'If List1.ListIndex > -1 Then
'    ParseDBsection Combo1.ListIndex, List1.ListIndex
'ElseIf List1.ListIndex < -1 Then
'    ParseDBsection Combo1.ListIndex, 32768 * 2& + List1.ListIndex
'End If
'End Sub

Private Function InitializeOffsets() As Boolean

' version 3+ byte definitions
' 5  << &H1 major version
' 8  << &H4 number of subs (1346)
' 12 << &H4 number of functions (4834)
' 16 << &H4 number of constants (52933)
' 20 << &H4 number of types (469)

' v2
' 24 << 1st delcaration text length
' 26 << begin 1st declaration text
' 26+lenText is next declaration length

' v3
' 24 << &H4 number of enumerations (4)
' 40 << 1st delcaration text length
' 42 << begin 1st declaration text
' 42+lenText is next declaration length

' ^^ each item (both versions) is always preceded by the length of the item (2 bytes)

Dim fnr As Integer
Dim byteOffset As Long, Looper As Long
Dim apiCount As Long, apiLen As Integer
Dim sectionIdx As Long
Dim apiBytes() As Byte
Dim fileVer As Byte

Const byteBase As Byte = 6 '5th byte used by Get# which is 1 bound vs 0 bound

ReDim apiBytes(0 To 23)
fnr = FreeFile()

Erase apiTotals
Erase apiOffsets

On Error GoTo BadFile
Open Text1.Text For Input Access Read As #fnr
Close #fnr

Open Text1.Text For Binary Access Read As #fnr
Get #fnr, byteBase, apiBytes()
    CopyMemory fileVer, apiBytes(0), &H1
    If fileVer = 0 Or fileVer > 3 Then ' check version
        MsgBox "Wrong File or Version - Compatible with version 2 & 3", vbInformation + vbOKOnly
        Close #fnr
        Exit Function
    End If
    CopyMemory apiTotals(apDeclarations), apiBytes(3), &H4
    CopyMemory apiTotals(apTypes), apiBytes(7), &H4
    apiTotals(apDeclarations) = apiTotals(apDeclarations) + apiTotals(apTypes)
    CopyMemory apiTotals(apConstants), apiBytes(11), &H4
    CopyMemory apiTotals(apTypes), apiBytes(15), &H4
    If fileVer > 2 Then CopyMemory apiTotals(apEnumerations), apiBytes(19), &H4

' now walk the rest of the file to find offsets
ReDim apiBytes(0 To 1)
byteOffset = Choose(fileVer, 25, 25, 41)
apiOffsets(apDeclarations) = byteOffset
For Looper = 1 To UBound(apiOffsets)
    For apiCount = 1 To apiTotals(sectionIdx)
        Get #fnr, byteOffset, apiBytes()
        CopyMemory apiLen, apiBytes(0), &H2
        byteOffset = byteOffset + apiLen + 2
    Next
    apiOffsets(Looper) = byteOffset
    Select Case Looper
    Case 1: ' declarations
        sectionIdx = apDeclarations
    Case 2, 3: ' types
        sectionIdx = apTypes
    Case 4, 5: ' constants
        sectionIdx = apConstants
    Case 6: ' enumerations
        sectionIdx = apEnumerations
    End Select
Next
Close #fnr
InitializeOffsets = True
Exit Function

BadFile:
MsgBox "Add the full path & file name to your .apv file in the text box, then open form", vbOKOnly
Unload Me
End Function


Private Sub GetAPIlisting(Section As Integer, Optional ByVal fromIndex As Long = 0, Optional ByVal toIndex As Long = -1)

Dim fnr As Integer
Dim byteOffset As Long
Dim apiCount As Long
Dim apiLen As Integer
Dim apiN(0 To 1) As Byte
Dim apiBytes() As Byte
Dim apiItem As Long

' adding 52K constants to listbox takes time. Use LockWindowUpdate to speed it up

'List1.Clear
Text2 = ""
If toIndex < fromIndex Then toIndex = apiTotals(Section)

byteOffset = apiOffsets(Section * 2)
fnr = FreeFile()
Open Text1.Text For Binary Access Read As #fnr
Get #fnr, byteOffset, apiN()
For apiCount = 1 To apiTotals(Section)
    CopyMemory apiLen, apiN(0), &H2
    If apiCount >= fromIndex Then
        ReDim apiBytes(0 To apiLen - 1)
        Get #fnr, , apiBytes
        mockListItems(apiItem) = StrConv(apiBytes, vbUnicode)
        apiItem = apiItem + 1
    'List1.AddItem StrConv(apiBytes, vbUnicode)
        If apiItem = mockListCount - 1 Then Exit For
    End If
    byteOffset = byteOffset + apiLen + 2
    Get #fnr, byteOffset, apiN
Next
Close #fnr

End Sub


Private Sub ParseDBsection(Section As apiSections, Index As Long, Optional ByVal byteOffset As Long)
' this is pretty fast as is, however, should you want to
' pass an optional offset where the walking starts, you can

Dim fnr As Integer
Dim apiCount As Long
Dim apiData As String
Dim apiLen As Integer
Dim apiFormat As String
Dim I As Integer
Dim apiBytes() As Byte, apiParts() As String

fnr = FreeFile()
Open Text1.Text For Binary Access Read As #fnr
    
ReDim apiBytes(0 To 1)
If byteOffset = 0 Then byteOffset = apiOffsets(Section * 2 + 1)
' walk section to find item number
For apiCount = 0 To Index
    Get #fnr, byteOffset, apiBytes()
    CopyMemory apiLen, apiBytes(0), &H2
    byteOffset = byteOffset + apiLen + 2
Next
ReDim apiBytes(0 To apiLen - 1)
Get #fnr, , apiBytes
Close #fnr

apiData = StrConv(apiBytes, vbUnicode)

Select Case Section
Case apDeclarations 'format declarations
    If Right$(apiData, 1) <> ")" Then
        apiData = "Declare Function " & mockListItems(mockListIndex + mockTopIndex) & " Lib " & apiData
    Else
        apiData = "Declare Sub " & mockListItems(mockListIndex + mockTopIndex) & " Lib " & apiData
    End If

    I = InStr(apiData, Chr$(34))
    If I Then
        If InStr(apiData, ".") = 0 Then ' add the .dll to statement
            I = InStr(I + 1, apiData, Chr$(34))
            apiData = Left$(apiData, I - 1) & ".dll" & Mid$(apiData, I)
        End If
        I = InStr(I + 1, apiData, "(")
        apiParts = Split(Mid$(apiData, I + 1), ",")
        apiData = Left$(apiData, I)
        For I = 0 To UBound(apiParts)
            Select Case Left$(LTrim(apiParts(I)) & ")", 1)
            Case ")"
            Case "?"
                apiParts(I) = Replace$(LTrim$(apiParts(I)), "?", "ByVal ")
            Case "~"
                apiParts(I) = Replace$(LTrim$(apiParts(I)), "~", "ByRef ")
            Case Else
                apiParts(I) = "ByRef " & Mid$(LTrim$(apiParts(I)), 1)
            End Select
        Next
        apiData = apiData & Join(apiParts, ", ")
        apiData = Replace$(apiData, "&", " As Long")
        apiData = Replace$(apiData, "$", " As String")
        apiData = Replace$(apiData, "%", " As Integer")
        'apiData = Replace$(apiData, "#", " As Double") '<< not currently used
        'apiData = Replace$(apiData, "!", " As Single") '<< not currently used
        'apiData = Replace$(apiData, "@", " As Currency") '<< not currently used
    Else
        apiData = ""
    End If
Case apTypes
    apiData = "Type " & mockListItems(mockListIndex + mockTopIndex) & vbCrLf & apiData & vbCrLf & "End Type"
Case apEnumerations
    apiData = "Enum " & mockListItems(mockListIndex + mockTopIndex) & vbCrLf & apiData & vbCrLf & "End Enum"
Case apConstants
    ' same routine the apiviewer must use as it reports the wrong Type: Long vs String for...
    '   Const wszNAMESEPARATORDEFAULT As **Long** = szNAMESEPARATORDEFAULT
    '   where szNAMESEPARATORDEFAULT is declared As **String** = "\n"
    ' The right way would be to look up the referenced constant to be absolutely sure
    
    If InStr(apiData, Chr$(34)) Then
        apiData = "Const " & mockListItems(mockListIndex + mockTopIndex) & " As String = " & apiData
    Else
        apiData = "Const " & mockListItems(mockListIndex + mockTopIndex) & " As Long = " & apiData
    End If
End Select
Text2.Text = apiData
End Sub

Private Function SearchAPIFile(Section As apiSections, Criteria As String, Optional ByVal byteOffset As Long)
' Should you want to pass an optional offset where the walking starts, you can

' example of doing an exact match/find (i.e., looking for a specific TYPE
' when the TYPE is referenced in a declaration statement)

Dim fnr As Integer
Dim apiCount As Long
Dim apiLen As Integer
Dim apiN(0 To 1) As Byte
Dim apiBytes() As Byte
Dim lRtn As Long

lRtn = -1
If byteOffset = 0 Then byteOffset = apiOffsets(Section * 2)
fnr = FreeFile()
Open Text1.Text For Binary Access Read As #fnr
Get #fnr, byteOffset, apiN()
For apiCount = 1 To apiTotals(Section)
    CopyMemory apiLen, apiN(0), &H2
    ReDim apiBytes(0 To apiLen - 1)
    Get #fnr, , apiBytes
    If StrComp(StrConv(apiBytes, vbUnicode), Criteria, vbTextCompare) = 0 Then
        lRtn = apiCount - 1
        Exit For
    End If
    Get #fnr, , apiN
Next
Close #fnr
SearchAPIFile = lRtn
End Function

Private Sub UpdateMockLB()
Dim x As Long, tRect As RECT
Dim f As Long

DrawText mockListBox.hdc, "Wy", 2, tRect, DT_CALCRECT Or DT_SINGLELINE
tRect.Right = mockListBox.Width / Screen.TwipsPerPixelX - lbScroll.Width \ Screen.TwipsPerPixelX
tRect.Left = 3
mockListBox.Cls
For x = 0 To mockListCount - 1
    If x = mockListIndex Then
        FillRect mockListBox.hdc, tRect, hBrushFill(1)
        SetTextColor mockListBox.hdc, vbWhite
    End If
    DrawText mockListBox.hdc, mockListItems(x), -1, tRect, DT_SINGLELINE Or DT_VCENTER
    OffsetRect tRect, 0, tRect.Bottom - tRect.Top
    If x + mockTopIndex = apiTotals(Combo1.ListIndex) Then Exit For
    If x = mockListIndex Then SetTextColor mockListBox.hdc, vbBlack
Next
For f = x To mockListCount - 1
    FillRect mockListBox.hdc, tRect, hBrushFill(0)
    OffsetRect tRect, 0, tRect.Bottom - tRect.Top
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
If hBrushFill(0) Then
    DeleteObject hBrushFill(0)
    DeleteObject hBrushFill(1)
End If
End Sub

Private Sub lbScroll_Change()
If Len(lbScroll.Tag) Then Exit Sub

If bScrolling Then
    mockTopIndex = lbScroll.Value * mockScrollRatio
Else
    mockTopIndex = lbScroll.Value
End If
GetAPIlisting Combo1.ListIndex, mockTopIndex, mockTopIndex + mockListCount - 1
UpdateMockLB
mockListBox.Refresh
If Not bScrolling Then
    lbScroll.Tag = "NoRecurse"
    lbScroll.Value = mockTopIndex / mockScrollRatio
    lbScroll.Tag = ""
Else
'    bScrolling = False
End If
End Sub

Private Sub lbScroll_Scroll()
bScrolling = True
Call lbScroll_Change
End Sub

Private Sub mockListBox_Resize()
Dim tRect As RECT
If hBrushFill(0) = 0 Then
    hBrushFill(0) = CreateSolidBrush(vbWhite)
    hBrushFill(1) = CreateSolidBrush(vbBlue)
End If
DrawText mockListBox.hdc, "Wy", 2, tRect, DT_CALCRECT Or DT_SINGLELINE
mockListCount = (mockListBox.Height \ Screen.TwipsPerPixelY) \ tRect.Bottom + 1
mockTopIndex = 0
mockListIndex = 9
lbScroll.Tag = "NoRecurse"
lbScroll.Value = 0
lbScroll.LargeChange = mockListCount
lbScroll.Tag = ""
ReDim mockListItems(0 To mockListCount - 1)
If apiTotals(Combo1.ListIndex) > 1000 Then
    mockScrollRatio = apiTotals(Combo1.ListIndex) / 1000
Else
    mockScrollRatio = 1
End If
GetAPIlisting Combo1.ListIndex, mockTopIndex, mockTopIndex + mockListCount - 1
UpdateMockLB
End Sub
