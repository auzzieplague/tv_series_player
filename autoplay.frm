VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tv Series Helper"
   ClientHeight    =   1530
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   3375
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox filmtitle 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   720
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   5520
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   615
      Left            =   6480
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Rename && Order"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2760
      Top             =   840
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   2880
      X2              =   3120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      X1              =   2880
      X2              =   3000
      Y1              =   960
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   3120
      X2              =   3000
      Y1              =   960
      Y2              =   1200
   End
   Begin VB.Menu mnu1 
      Caption         =   "Menu"
      Begin VB.Menu mnu2 
         Caption         =   "Reset Count"
      End
      Begin VB.Menu mnu3 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute _
    Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    Function resetCount()
     Kill App.Path + "\count.txt"
    Form1.Caption = "Upto File " + str(getCount() + 1)
    End
    End Function
    Function removeExtraSpaces(txt As String) As String
        Dim i As Integer
        i = 1
        On Error GoTo handler:
        
        Do While (i < Len(txt))
            Do While (Mid$(txt, i, 1) <> " ")
                removeExtraSpaces = removeExtraSpaces + Mid$(txt, i, 1)
                i = i + 1
            Loop
            removeExtraSpaces = removeExtraSpaces + " "
            Do While (Mid$(txt, i, 1) = " ")
                i = i + 1
            Loop
        Loop
        
        Exit Function
        
handler::
        'expected error from loop running past end of string
    End Function
    
    Function captialiseText(txt As String) As String
        Dim arr
        txt = LCase$(txt)
        Dim i As Integer
        
        arr = Split(txt, " ")
        For i = 0 To UBound(arr) - 1
            If Len(arr(i)) > 0 Then Mid$(arr(i), 1, 1) = UCase$(Left$(arr(i), 1))
        Next i
        
        captialiseText = Join(arr, " ")
    End Function
    Function getSeriesName() As String
    Dim i, c As Integer
    Dim parts
    Dim whole As String
    
    'remove file extension
    whole = Left$(File1.List(0), Len(File1.List(0)) - 4)
    'join and split again to seperate using different delimters
    whole = Join(Split(whole, "."), " ")
    whole = Join(Split(whole, "_"), " ")
    whole = Join(Split(whole, "-"), " ")
    
    'remove numbers from title
    For i = 1 To Len(whole)
        If isNumber(Mid$(whole, i, 1)) Then
        Mid$(whole, i, 1) = " "
        End If
    Next i
    
    'work in lowercase
    whole = LCase$(whole)
    
    'because numbers have been replaces with spaces, words can be identified with using spacing pattern
    whole = Join(Split(whole, " episode "), " ")
    whole = Join(Split(whole, " season "), " ")
    whole = Join(Split(whole, " s "), " ")      'lonely s from s01
    whole = Join(Split(whole, " se "), " ")      'lonely s from s01
    whole = Join(Split(whole, " ep "), " ")     'lonely ep from ep01
    whole = Join(Split(whole, " e "), " ")      'lonely e from s01e01 etc
    
    'clean up gaps
    whole = removeExtraSpaces(whole)
    
    'last split into sections
    parts = Split(whole, " ")
    
    'find common words in all filenames
    For i = 1 To File1.ListCount - 1
        For c = 0 To UBound(parts)
            If InStr(LCase(File1.List(i)), parts(c)) = 0 Then
            'if parts are not found in file part[] to empty
            parts(c) = ""
            End If
        Next c
    Next i
    'join parts again to get full title with common words
    getSeriesName = captialiseText(Join(parts, " "))
    
    
    End Function
    
    Function isNumber(char As String) As Boolean
        isNumber = False
        On Error GoTo handler:
        Dim ascii As Integer
        
        ascii = Asc(Left$(char, 1))
        If ascii > 47 And ascii < 58 Then
                isNumber = True
        End If
        Exit Function
handler::
        'not a number
    End Function
    Function countSetsOfNumbers(txt As String) As Integer
       Dim i As Integer
        On Error GoTo handler:
        i = 0
        
        countSetsOfNumbers = 0
        Do While i < Len(txt)
            Do
                i = i + 1
            Loop Until (isNumber(Mid$(txt, i, 1)))
            countSetsOfNumbers = countSetsOfNumbers + 1
                'skip digits
            Do While (isNumber(Mid$(txt, i, 1)))
                i = i + 1
            Loop
        Loop
        Exit Function
handler::
        'expected error
    End Function
    Function getNumbersSet(txt As String, count As Integer) As String
        Dim i As Integer
        Dim c As Integer
        On Error GoTo handler:
        getNumbersSet = ""
        i = 0
        c = 0
        
        Do While c < count
            Do
                i = i + 1
            Loop Until (isNumber(Mid$(txt, i, 1)))
            c = c + 1
            If c < count Then
                'skip digits
                Do While (isNumber(Mid$(txt, i, 1)))
                    i = i + 1
                Loop
            End If
        Loop
        
        Do While (isNumber(Mid$(txt, i, 1)))
            getNumbersSet = getNumbersSet + Mid$(txt, i, 1)
            i = i + 1
        Loop
        Exit Function
handler::
        'expected error
    End Function
Function getFileExtension(txt As String) As String
    getFileExtension = Right$(txt, 3)
End Function
Function reformatFilenames()
    'find largest number of consecutive digits and reformat, padding filenumbers with 0's
    Dim digitCount As Integer
    digitCount = findMaxConsecutiveNumbers
    Dim i As Integer
    Dim setsCount As Integer
    Dim newName As String
    Dim title, formattedNumbers As String
    title = filmtitle.Text
    Dim response As Integer
    
    response = MsgBox("Would you like me to rename and number your " + filmtitle.Text + " files?", vbYesNo, "Confirm Rename")
    If response = vbNo Then Exit Function
    
    For i = 0 To File1.ListCount - 1
        
        setsCount = countSetsOfNumbers(File1.List(i)) 'they should all have same but just incase
        formattedNumbers = padNumbers(File1.List(i), digitCount)
        Select Case setsCount
        Case 1
            newName = title + " E" + getNumbersSet(formattedNumbers, 1) + "." + getFileExtension(File1.List(i))
        Case Is > 1
            newName = title + " S" + getNumbersSet(formattedNumbers, 1) + "E" + getNumbersSet(formattedNumbers, 2) + "." + getFileExtension(File1.List(i))
        End Select
        'List1.AddItem (newName)
        'List1.AddItem (getNumbersSet(formattedNumbers, 2))
        Name (File1.Path + "\" + File1.List(i)) As File1.Path + "\" + newName
    Next i
    
    'refresh filelistbox
    File1.Path = "c:"
    File1.Path = App.Path
    
    MsgBox "Files have been processed", vbInformation, "all done!"
End Function
Function padNumbers(str As String, digits As Integer) As String
    Dim i, n, ascii As Integer
    Dim curlen As Integer
    Dim build As String
    Dim final As String
    str = str + " " 'pad string for beter handling of loop
    For n = 1 To Len(str)
         curlen = 0
         build = ""
         Do While (isNumber(Mid$(str, n, 1)) And n <= Len(str))
             build = build + (Mid$(str, n, 1))
             curlen = curlen + 1
             n = n + 1
         Loop
         If curlen < digits And curlen > 0 Then
            Do While curlen < digits
                build = "0" + build
                curlen = curlen + 1
            Loop
         End If
         final = final + build + (Mid$(str, n, 1))
     Next n
     padNumbers = final
End Function
Function findMaxConsecutiveNumbers() As Integer
    Dim i, n, ascii As Integer
    Dim curlen As Integer
    findMaxConsecutiveNumbers = 0
    For i = 0 To File1.ListCount - 1
        For n = 1 To Len(File1.List(i))
            curlen = 0
            Do While isNumber(Mid$(File1.List(i), n, 1))
                curlen = curlen + 1
                n = n + 1
            Loop
            If curlen > findMaxConsecutiveNumbers Then findMaxConsecutiveNumbers = curlen
        Next n
    Next i
End Function
Private Sub Command1_Click()
    If getCount() > File1.ListCount - 1 Then
        MsgBox "Reached end of list"
    Else
        File1.ListIndex = getCount()
    End If
    runFile File1.FileName
    increaseCount
    End
End Sub

Private Sub Form_Load()

    File1.Path = App.Path
    File1.Pattern = "*.avi;*.mkv;*.mp4"
    filmtitle = getSeriesName
    Form1.Left = Screen.Width / 2 - Form1.Width / 2
    Form1.Top = Screen.Height / 2 - Form1.Height / 2
    Form1.Caption = "[" + Trim(str(getCount() + 1)) + "] " + filmtitle
End Sub

Function getCount() As Integer
    Dim count As Variant
    On Error GoTo handler:
    Open App.Path + "\count.txt" For Input As #1
        Line Input #1, count
    Close #1
    getCount = count
Exit Function
handler::
    Close #1
    getCount = 0
End Function
Function runFile(sFile As String)
    
    Dim sCommand As String
    Dim sWorkDir As String
    
    sCommand = vbNullString         'Command line parameters
    sWorkDir = App.Path

    ShellExecute hwnd, "open", sFile, sCommand, sWorkDir, 1

End Function

Function increaseCount()
    Dim count As Integer
    count = getCount()
    Open App.Path + "\count.txt" For Output As #1
        Write #1, count + 1
    Close #1
End Function

Private Sub Image1_Click()
    reformatFilenames
End Sub

Private Sub mnu2_Click()
    resetCount
End Sub

Private Sub mnu3_Click()
    frmAbout.Show
End Sub
