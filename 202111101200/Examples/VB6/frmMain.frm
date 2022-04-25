VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   Caption         =   "Neural Network Tester"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9435
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   9  'Size W E
   ScaleHeight     =   6435
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   3720
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   4830
      Visible         =   0   'False
      Width           =   705
   End
   Begin MSComDlg.CommonDialog cdi1 
      Left            =   2370
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picOutputs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   6990
      MousePointer    =   1  'Arrow
      ScaleHeight     =   4365
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.PictureBox picInputs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   30
      MousePointer    =   1  'Arrow
      ScaleHeight     =   4365
      ScaleWidth      =   6885
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   6915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FORM_CAPTION = "Neural Network Viewer"
Const FORM_CAPTION_SEPARATOR = " - "

Private mSngPctIO As Single

Private mIntICols As Integer 'number of input columns
Private mIntIRows As Integer 'number of input rows

Private mIntOCols As Integer 'number of output columns
Private mIntORows As Integer 'number of output columns

Private mPages() As tIOPage
Private mLngPages As Long 'total number of pages
Private mLngPage_Current As Long 'current page

Private mBolTraining As Boolean
Private mBolActivated As Boolean

Private mBolThinking As Boolean 'showing outputs (false) or thinking the outputs (true)

Private Const MAX_NUMBER_OF_NEURONS = 200000

Private Const PERCENTAGE_OF_ACTIVATION_TO_CONSIDER_ACTIVE = 0.2

Private mLngColorGrey As Long

Private mSngInk As Single

'to find out height of title bar
Private Const SM_CYCAPTION As Long = 4
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Sub Form_Activate()
  
  If Not mBolActivated Then
    MousePointer = 11
    mBolActivated = InitializeAutomatically()
    If mBolActivated Then Call ShowHelp Else Unload Me
    MousePointer = 9 'puts the icon to reescale
  End If

End Sub

Private Function InitializeWithANewNumberOfLayers(ByVal pIntICols As Integer, _
                                                  ByVal pIntIRows As Integer, _
                                                  ByVal pIntOCols As Integer, _
                                                  ByVal pIntORows As Integer, _
                                                  ByVal pBolSilent As Boolean) As Boolean
  Dim lIntNumberOfLayers As Integer
  
  If pBolSilent Then
    lIntNumberOfLayers = pIntORows + 1
  Else
    lIntNumberOfLayers = Val(InputBox("Number of layers?", "Number of layers", pIntORows + 1))
  End If
  
  If lIntNumberOfLayers >= 2 Then
    InitializeWithANewNumberOfLayers = NeuralNetworkCreate(pIntICols, pIntIRows, pIntOCols, pIntORows, lIntNumberOfLayers)
  End If

End Function

Private Function InitializeAutomatically() As Boolean

  If InitializeWithANewNumberOfLayers(mIntICols, mIntIRows, mIntOCols, mIntORows, True) Then
    Call InitializeInternals
    InitializeAutomatically = True
  End If

End Function

Private Function Initialize() As Boolean
  Dim lIntICols As Integer
  Dim lIntIRows As Integer
  Dim lIntOCols As Integer
  Dim lIntORows As Integer
  Dim lIntNumberOfLayers As Integer
  
  'of this rows and cols in inputs
  lIntICols = Val(InputBox("Number of columns in inputs?", "Input columns", mIntICols))
  If lIntICols <> 0 Then
    lIntIRows = Val(InputBox("Number of rows in inputs?", "Input rows", mIntIRows))
    If lIntIRows <> 0 Then
      'and outputs
      lIntOCols = Val(InputBox("Number of columns in outputs?", "Outputs columns", mIntOCols))
      If lIntOCols <> 0 Then
        lIntORows = Val(InputBox("Number of rows in outputs?", "Outputs rows", mIntORows))
        If lIntORows <> 0 Then
          If InitializeWithANewNumberOfLayers(lIntICols, lIntIRows, lIntOCols, lIntORows, False) Then
            mIntICols = lIntICols
            mIntIRows = lIntIRows
            mIntOCols = lIntOCols
            mIntORows = lIntORows
            Call InitializeInternals
            Initialize = True
          End If
        End If
      End If
    End If
  End If

End Function

Private Sub InitializeInternals()

  mLngPages = 1 'first page
  mLngPage_Current = 0
  ReDim mPages(mLngPages - 1) As tIOPage 'adjusts memory
  mSngPctIO = picInputs.Width / Me.Width
  picInputs.Visible = True
  picOutputs.Visible = True
  Call IOReserveMemoryForThisPage
  Call UpdateForm
  Call CaptionUpdate((mLngPage_Current + 1) & "/" & mLngPages)

End Sub

Private Function PagesSave() As Boolean
  
  With cdi1
    .Filter = "NNViewer (internally is a CSV)(*.nnv)|*.nnv"
    .InitDir = App.Path
    .FileName = vbNullString
    .ShowSave
    If LenB(.FileName) <> 0 Then
      Me.MousePointer = 11
      PagesSave = PagesSaveToFile(.FileName)
      Me.MousePointer = 9
    End If
  End With

End Function

Private Function PagesLoad() As Boolean
  
  With cdi1
    .Filter = "NNViewer (internally is a CSV)(*.nnv)|*.nnv|Icon (*.ico)|*.ico"
    .InitDir = App.Path
    .FileName = vbNullString
    .ShowOpen
    If LenB(.FileName) <> 0 Then
      Me.MousePointer = 11
      Select Case LCase$(Right$(.FileName, 4))
      Case ".nnv": PagesLoad = PagesLoadFromFile_NNV(.FileName)
      Case ".ico": PagesLoad = PagesLoadFromFile_ICO(.FileName)
      End Select
      Me.MousePointer = 9
    End If
  End With

End Function

Private Function PagesLoadFromFile_NNV(ByVal pStrFileName As String) As Boolean
  Dim lLngFile As Long
  Dim i As Long
  Dim j As Long
  Dim p As Long
  Dim lStrTmp As String
  Dim lIntICols As Integer
  Dim lIntIRows As Integer
  Dim lIntOCols As Integer
  Dim lIntORows As Integer
  Dim lLngPages As Long
  Dim lPieces() As String
  Dim lBolError As Boolean
  Dim lPages() As tIOPage
  
  #If vbDebug = 0 Then
    On Error GoTo Errores
  #End If
  
  'default
  lBolError = True

  lLngFile = FreeFile()
  If lLngFile <> 0 Then
    Open pStrFileName For Input As #lLngFile
    Line Input #lLngFile, lStrTmp
    lPieces = Split(lStrTmp, ";")
    If SafeUbound(lPieces) = 4 Then
      lIntICols = lPieces(0)
      lIntIRows = lPieces(1)
      lIntOCols = lPieces(2)
      lIntORows = lPieces(3)
      If InitializeWithANewNumberOfLayers(lIntICols, lIntIRows, lIntOCols, lIntORows, False) Then
        lLngPages = lPieces(4)
        ReDim lPages(lLngPages - 1) As tIOPage
        lBolError = False
        For p = 0 To lLngPages - 1
          With lPages(p)
            ReDim .SngInputs(lIntICols * lIntIRows - 1) As Single
            ReDim .SngOutputs(lIntOCols * lIntORows - 1) As Single
          End With
          For i = 0 To lIntIRows - 1
            Line Input #lLngFile, lStrTmp
            lPieces = Split(lStrTmp, ";")
            If SafeUbound(lPieces) >= lIntICols Then
              For j = 0 To lIntICols - 1
                lPages(p).SngInputs(i * lIntICols + j) = lPieces(j)
              Next j
            Else
              MsgBox "Incorrect format in page " & p & " input row " & i, vbCritical, App.Title
              lBolError = True
              GoTo CloseAndExit
            End If
          Next i
          For i = 0 To lIntORows - 1
            Line Input #lLngFile, lStrTmp
            lPieces = Split(lStrTmp, ";")
            If SafeUbound(lPieces) >= lIntOCols Then
              For j = 0 To lIntOCols - 1
                lPages(p).SngOutputs(i * lIntOCols + j) = lPieces(j)
              Next j
            Else
              MsgBox "Incorrect format in page " & p & " output row " & i, vbCritical, App.Title
              lBolError = True
              GoTo CloseAndExit
            End If
          Next i
        Next p
      End If
    Else
      MsgBox "Incorrect format: header should contain the number of input columns, input rows, output columns and output rows.", vbCritical, App.Title
    End If
CloseAndExit:
    Close #lLngFile
  End If
  
  'if all went ok
  If Not lBolError Then
    'substitute current pages
    mIntICols = lIntICols
    mIntIRows = lIntIRows
    mIntOCols = lIntOCols
    mIntORows = lIntORows
    mLngPages = lLngPages
    Erase mPages
    mPages = lPages
    mLngPage_Current = 0
    PagesLoadFromFile_NNV = True 'success
  End If
  
Fin:
  Exit Function
  
Errores:
  MsgBox "Error: " & Err.Number & "-" & Err.Description, vbCritical, App.Title
  Resume Fin
  
End Function

Private Function PagesLoadFromFile_ICO(ByVal pStrFileName As String) As Boolean
  Dim lIntCols As Integer
  Dim lIntRows As Integer
  Dim i As Long
  Dim j As Long
  Dim lBolTmp As Boolean

  #If vbDebug = 0 Then
    On Error GoTo Errores
  #End If
  
  'clears inputs
  For i = 0 To mIntICols * mIntIRows - 1
    mPages(mLngPage_Current).SngInputs(i) = 0
  Next i
  
  'puts picTmp out of sight
  picTmp.Visible = False
  picTmp.Left = Me.Width * 1.5
  picTmp.Visible = True
        
  picTmp.Picture = LoadPicture(pStrFileName)
  lIntCols = picTmp.Width / Screen.TwipsPerPixelX - 2
  lIntRows = picTmp.Height / Screen.TwipsPerPixelY - 2
  
  If lIntCols <> mIntICols Or lIntRows <> mIntIRows Then
    If MsgBox("Current inputs grid rows (" & mIntIRows & ") or columns (" & mIntICols & ") do not match the icon's rows (" & lIntRows & ") or columns (" & lIntCols & "). Would you like to adjust grid and neural network?", vbQuestion + vbYesNo, App.Title) = vbYes Then
      If InitializeWithANewNumberOfLayers(lIntCols, lIntRows, mIntOCols, mIntORows, False) Then
        mIntICols = lIntCols
        mIntIRows = lIntRows
        Call InitializeInternals
        lBolTmp = True
      End If
    Else
      Exit Function
    End If
  Else
    lBolTmp = True
  End If
  
  If lBolTmp Then
    For i = 0 To picTmp.ScaleHeight - 1
      For j = 0 To picTmp.ScaleWidth - 1
        mPages(mLngPage_Current).SngInputs(i * mIntICols + j) = 1 - picTmp.Point(j, i) / RGB(255, 255, 255)
      Next j
    Next i
    Call PaintIOs(picInputs, mIntICols, mIntIRows, mPages(mLngPage_Current).SngInputs)
    picTmp.Visible = False
  End If
  
Fin:
  Exit Function
  
Errores:
  MsgBox "Error: " & Err.Number & "-" & Err.Description, vbCritical, App.Title
  Resume Fin

End Function

Private Function PagesSaveToFile(ByVal pStrFileName As String) As Boolean
  Dim lLngFile As Long
  Dim i As Long
  Dim j As Long
  Dim p As Long
  Dim lStrTmp As String
  
  #If vbDebug = 0 Then
    On Error GoTo Errores
  #End If
  
  lLngFile = FreeFile()
  If lLngFile <> 0 Then
    Open pStrFileName For Output As #lLngFile
    Print #lLngFile, mIntICols & ";" & mIntIRows & ";" & mIntOCols & ";" & mIntORows & ";" & mLngPages
    For p = 0 To mLngPages - 1
      For i = 0 To mIntIRows - 1
        lStrTmp = vbNullString
        For j = 0 To mIntICols - 1
          lStrTmp = lStrTmp & mPages(p).SngInputs(i * mIntICols + j) & ";"
        Next j
        Print #lLngFile, lStrTmp
      Next i
      For i = 0 To mIntORows - 1
        lStrTmp = vbNullString
        For j = 0 To mIntOCols - 1
          lStrTmp = lStrTmp & mPages(p).SngOutputs(i * mIntOCols + j) & ";"
        Next j
        Print #lLngFile, lStrTmp
      Next i
    Next p
    Close #lLngFile
    PagesSaveToFile = True
  End If
  
Fin:
  Exit Function
  
Errores:
  MsgBox "Error: " & Err.Number & "-" & Err.Description, vbCritical, App.Title
  Resume Fin

End Function

Private Function NeuralNetworkCreate(ByVal pIntICols As Integer, _
                                     ByVal pIntIRows As Integer, _
                                     ByVal pIntOCols As Integer, _
                                     ByVal pIntORows As Integer, _
                                     ByVal pIntNumberOfLayers As Integer) As Boolean
  Dim lStrTmp As String

  Call CaptionUpdate("Creating neural network...")
  
  If CreateNeuralNetwork(MAX_NUMBER_OF_NEURONS, pIntICols * pIntIRows, pIntOCols * pIntORows, 0, pIntNumberOfLayers, pIntNumberOfLayers, lStrTmp) Then
    Call NetInitialize(0)
    MsgBox "Neural network was created successfully with nodes: " & vbCrLf & vbCrLf & lStrTmp, vbInformation, App.Title
    Call CaptionUpdate((mLngPage_Current + 1) & "/" & mLngPages)
    NeuralNetworkCreate = True 'success
  Else
    MsgBox "Error creating neural network.", vbCritical, App.Title
  End If

End Function

Private Sub ShowHelp()

  MsgBox "Mouse click to draw. Type: " & vbCrLf & vbCrLf & _
         "CTRL or SHIFT: switches from drawing or erasing." & vbCrLf & _
         "k: changes the ink intensity." & vbCrLf & _
         "i: initializes everything." & vbCrLf & _
         "y: initializes with a different number of layers." & vbCrLf & _
         "c: clears current page." & vbCrLf & _
         "n: creates new page." & vbCrLf & _
         "d: deletes current page." & vbCrLf & _
         "t: learns current page." & vbCrLf & _
         "a: learns all pages." & vbCrLf & _
         "ESC: stops learning." & vbCrLf & _
         "t: thinks to show outputs." & vbCrLf & _
         "v: verifies pages versus thought." & vbCrLf & _
         "Left & Right arrows, Start, End: moves between pages." & vbCrLf & _
         "s: saves pages." & vbCrLf & _
         "o: loads pages." & vbCrLf & _
         "m: shows this help.", vbInformation, App.Title
         
End Sub

Private Sub InkSelect()
  Dim lStrTmp As String
  
  lStrTmp = InputBox("Ink intensity (a value between 0 and 1)?", "Ink intensity", 1)
  If LenB(lStrTmp) <> 0 Then
    If IsNumeric(lStrTmp) Then
      If CSng(lStrTmp) > 1 Then
        mSngInk = 1
      ElseIf CSng(lStrTmp) <= 0 Then
        mSngInk = 0
      Else
        mSngInk = CSng(lStrTmp)
      End If
    End If
  End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim lBolUpdate As Boolean

  'Debug.Print KeyCode
  
  MousePointer = 11
  
  Select Case KeyCode
  'CTRL or SHIFT
  Case 17, 16:
    If mSngInk >= 0.5 Then mSngInk = 0 Else mSngInk = 1
  Case 27:
    'stop studying
    mBolTraining = False
    lBolUpdate = True
  'k: select ink
  Case 75:
    Call InkSelect
  'i: initialize
  Case 73:
    Call Initialize
  'y: initialize number of layers
  Case 89:
    Call InitializeWithANewNumberOfLayers(mIntICols, mIntIRows, mIntOCols, mIntORows, False)
  'c: clear
  Case 67:
    Call IOClear
  'n: new
  Case 78:
    Call MemorizeIOAndCreateNewPaper
    Call CaptionUpdate((mLngPage_Current + 1) & "/" & mLngPages)
  'Key to go to 1st page
  Case 36:
    mLngPage_Current = 0
    lBolUpdate = True
  'key to go to last page
  Case 35:
    mLngPage_Current = mLngPages - 1
    lBolUpdate = True
  'Right arrow
  Case 39:
    If mLngPage_Current < mLngPages - 1 Then
      mLngPage_Current = mLngPage_Current + 1
      lBolUpdate = True
    End If
  'left arrow
  Case 37:
    If mLngPage_Current > 0 Then
      mLngPage_Current = mLngPage_Current - 1
      lBolUpdate = True
    End If
  'a: learn set
  Case 65:
    Call CaptionUpdate((mLngPage_Current + 1) & "/" & mLngPages & FORM_CAPTION_SEPARATOR & " Studying"):
    Call LearnSet
    Call PaintIOs(picInputs, mIntICols, mIntIRows, mPages(mLngPage_Current).SngInputs)
    Call PaintIOs(picOutputs, mIntOCols, mIntORows, Outputs())
  'l: learn page
  Case 76:
    Call LearnPage
    Call CaptionUpdate((mLngPage_Current + 1) & "/" & mLngPages & FORM_CAPTION_SEPARATOR & " Learned")
  't: think and say what are these inputs
  Case 84:
    mBolThinking = Not mBolThinking
    lBolUpdate = True
  'd: delete current page
  Case 68:
    If mLngPages > 1 Then
      Call PageDelete
      lBolUpdate = True
    End If
  'h:
  Case 72
    MsgBox "Hardware id of this device is (you can paste it as has been copied to your clipboard): " & MyHardwareIdAndIntoClipBoard(), vbInformation
  'm:
  Case 77:
    Call ShowHelp
  'r: random
  Case 82:
    Call PagesCreateRandom
  'v: verify matching of inputs & outputs
  Case 86:
    Call PagesVerify
  's: save all pages
  Case 83:
    If PagesSave() Then Call CaptionUpdate((mLngPage_Current + 1) & "/" & mLngPages & FORM_CAPTION_SEPARATOR & " Saved to file")
  'o: load all pages
  Case 79:
    lBolUpdate = PagesLoad()
  End Select
  
  If lBolUpdate Then Call PageRefresh
  
  MousePointer = 9

End Sub

Private Sub PageRefresh()

  Call PaintIOs(picInputs, mIntICols, mIntIRows, mPages(mLngPage_Current).SngInputs)
  Call PaintIOs(picOutputs, mIntOCols, mIntORows, Outputs())
  Call CaptionUpdate((mLngPage_Current + 1) & "/" & mLngPages)

End Sub

Private Sub PagesVerifyBasedOnHorizontalVisualDivision()
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim lLngTotalGood As Long
  Dim lLngPage_Current_Prev As Long
  Dim lBolThisPageIsCorrect As Boolean
  Dim lSngOutputs() As Single
  Dim lIntTmp As Integer
  Dim lSngValueToConsiderActive As Single
  Dim lStrTmp As String
  
  #If vbDebug = 0 Then
    On Error GoTo Errores
  #End If
    
  lStrTmp = InputBox("Value to consider active?", "Value for active", 0.8)
  If LenB(lStrTmp) <> 0 Then
    lSngValueToConsiderActive = lStrTmp
    If lSngValueToConsiderActive <> 0 Then
      lStrTmp = vbNullString
      lLngPage_Current_Prev = mLngPage_Current
      lIntTmp = (mIntICols * mIntIRows / mIntORows)
      For i = 0 To mLngPages - 1
        mLngPage_Current = i
        lSngOutputs = Outputs()
        For j = 0 To mIntORows - 1
          'if output is not activated, then all corresponding inputs must also be not activated
          If lSngOutputs(j) < lSngValueToConsiderActive Then
            'default
            lBolThisPageIsCorrect = True
            For k = j * lIntTmp To (j + 1) * lIntTmp - 1
              If mPages(i).SngInputs(k) >= lSngValueToConsiderActive Then
                'this page is not correct
                lBolThisPageIsCorrect = False
                Exit For
              End If
            Next k
          'if output is activated, at least 1 input must also be activated
          Else
            'default
            lBolThisPageIsCorrect = False
            For k = j * lIntTmp To (j + 1) * lIntTmp - 1
              If mPages(i).SngInputs(k) >= lSngValueToConsiderActive Then
                lBolThisPageIsCorrect = True
                Exit For
              End If
            Next k
          End If
          If Not lBolThisPageIsCorrect Then Exit For
        Next j
        If lBolThisPageIsCorrect Then lLngTotalGood = lLngTotalGood + 1 Else lStrTmp = lStrTmp & i + 1 & ","
      Next i
      mLngPage_Current = lLngPage_Current_Prev
      If LenB(lStrTmp) <> 0 Then lStrTmp = ", errors in pages: " & Left$(lStrTmp, Len(lStrTmp) - 1)
      MsgBox "Percentage of matches is " & FormatNumber(lLngTotalGood / mLngPages * 100, 2, True, False, False) & "%" & lStrTmp, vbInformation, App.Title
    End If
  End If
  
Fin:
  Exit Sub
  
Errores:
  MsgBox "Error: " & Err.Number & "-" & Err.Description, vbCritical, App.Title
  Resume Fin
  
End Sub

Private Sub PagesVerify()
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim lLngTotalGood As Long
  Dim lLngPage_Current_Prev As Long
  Dim lBolThisPageIsCorrect As Boolean
  Dim lSngOutputs() As Single
  Dim lIntTmp As Integer
  Dim lSngValueToConsiderActive As Single
  Dim lStrTmp As String
  
  #If vbDebug = 0 Then
    On Error GoTo Errores
  #End If
    
  lStrTmp = InputBox("Value to consider active?", "Value for active", 0.8)
  If LenB(lStrTmp) <> 0 Then
    lSngValueToConsiderActive = lStrTmp
    If lSngValueToConsiderActive <> 0 Then
      lStrTmp = vbNullString
      lLngPage_Current_Prev = mLngPage_Current
      lIntTmp = (mIntICols * mIntIRows / mIntORows)
      For i = 0 To mLngPages - 1
        mLngPage_Current = i
        lSngOutputs = Outputs()
        lBolThisPageIsCorrect = True 'default
        For j = 0 To mIntORows * mIntOCols - 1
          If lSngOutputs(j) >= lSngValueToConsiderActive Then
            If mPages(i).SngOutputs(j) < lSngValueToConsiderActive Then
              lBolThisPageIsCorrect = False
              Exit For
            End If
          ElseIf mPages(i).SngOutputs(j) >= lSngValueToConsiderActive Then
            lBolThisPageIsCorrect = False
            Exit For
          End If
        Next j
        If lBolThisPageIsCorrect Then lLngTotalGood = lLngTotalGood + 1 Else lStrTmp = lStrTmp & i + 1 & ","
      Next i
      mLngPage_Current = lLngPage_Current_Prev
      If LenB(lStrTmp) <> 0 Then lStrTmp = ", errors in pages: " & Left$(lStrTmp, Len(lStrTmp) - 1)
      MsgBox "Percentage of matches is " & FormatNumber(lLngTotalGood / mLngPages * 100, 2, True, False, False) & "%" & lStrTmp, vbInformation, App.Title
    End If
  End If
  
Fin:
  Exit Sub
  
Errores:
  MsgBox "Error: " & Err.Number & "-" & Err.Description, vbCritical, App.Title
  Resume Fin
  
End Sub

Private Sub PagesCreateRandom()
  Dim lIntTmp1 As Integer
  Dim lIntTmp2 As Integer
  Dim lSngTmp1 As Single
  Dim lIntTmp3 As Integer
  Dim i As Long
  Dim j As Long
  
  If mIntOCols <> 1 Then
    MsgBox "Random pages can only be created when there is only 1 column of outputs.", vbCritical, App.Title
  Else
    lIntTmp1 = Val(InputBox("How many random pages do you want to create?", "Number of random pages", 100))
    If lIntTmp1 > 0 Then
      lSngTmp1 = Val(InputBox("Error rate in percentage?", "Error rate", 0))
      MousePointer = 11
      Randomize
      For i = 1 To lIntTmp1
        mLngPages = mLngPages + 1
        Call IOReserveMemoryForThisPage
        mLngPage_Current = mLngPages - 1
        'puts random inputs and outputs, but following visual distribution
        lIntTmp1 = Int(Rnd() * mIntORows)
        lIntTmp2 = (mIntICols * mIntIRows - 1) / mIntORows
        'generates inputs
        For j = 1 To 30
          Do
            lIntTmp3 = lIntTmp2 * lIntTmp1 + Int(Rnd() * lIntTmp2)
          Loop Until lIntTmp3 <= SafeUbound(mPages(mLngPage_Current).SngInputs)
          mPages(mLngPage_Current).SngInputs(lIntTmp3) = 1
        Next j
        If lSngTmp1 = 0 Then 'if you want to see how it recognizes certain paterns, for example the first row, add: Or lIntTmp1 = 0
          mPages(mLngPage_Current).SngOutputs(lIntTmp1) = 1
        'if no random error
        ElseIf Rnd() * 100 >= lSngTmp1 Then
          mPages(mLngPage_Current).SngOutputs(lIntTmp1) = 1
        'if random error
        Else
          mPages(mLngPage_Current).SngOutputs(Int(Rnd() * mIntORows)) = 1
        End If
        Call PaintIOs(picInputs, mIntICols, mIntIRows, mPages(mLngPage_Current).SngInputs)
        Call PaintIOs(picOutputs, mIntOCols, mIntORows, mPages(mLngPage_Current).SngOutputs)
        Call CaptionUpdate((mLngPage_Current + 1) & "/" & mLngPages)
      Next i
      MousePointer = 9
    End If
  End If
  
End Sub

Private Sub PageDelete()
  Dim i As Long
  
  If mLngPage_Current < mLngPages - 1 Then
    For i = mLngPage_Current To mLngPages - 2
      mPages(i).SngInputs = mPages(i + 1).SngInputs
      mPages(i).SngOutputs = mPages(i + 1).SngOutputs
    Next i
  End If
  
  mLngPages = mLngPages - 1
  If mLngPage_Current + 1 > mLngPages Then mLngPage_Current = mLngPages - 1
  ReDim Preserve mPages(mLngPages) As tIOPage
  
End Sub

Private Function Outputs() As Single()
  Dim i As Long
  Dim j As Long
  Dim lSngOutputs() As Single
  Dim lLngCount As Long
  
  If mBolThinking Then
    ReDim lSngOutputs(mIntOCols * mIntORows - 1) As Single
    Call PutInputs(mLngPage_Current)
    For i = 0 To mIntORows - 1
      For j = 0 To mIntOCols - 1
        lSngOutputs(i * mIntOCols + j) = NetOutputGet(lLngCount, 0)
        lLngCount = lLngCount + 1
      Next j
    Next i
    Outputs = lSngOutputs
  Else
    Outputs = mPages(mLngPage_Current).SngOutputs
  End If
      
End Function

Private Sub LearnSet()
  Dim lDatTmp As Date
  Dim lLngTmp As Long
  Dim i As Long
  Dim j As Long
  Dim lSngPercentOk As Single
  Dim lSngPercentMax As Single
  Dim lSngPercentOk_Prev As Single
  Dim lLngCycleWhereMaxHappened As Long
  Dim lDatStart As Date
  Dim lSngPercentageTarget As Single
    
  mBolTraining = False 'stops previous trainings
  
  'calculates time for 1 training
  lDatTmp = Now()
  Call LearnPage
  lLngTmp = DateDiff("s", lDatTmp, Now())
  lLngTmp = Val(InputBox("Training 1 item took " & lLngTmp & " seconds. How many iterations do you want to do now?", "Iterations", 0))
  If lLngTmp <> 0 Then
    lSngPercentageTarget = Val(InputBox("Target percentage of success?", "Target percentage", 95))
    If lSngPercentageTarget <> 0 Then
      lSngPercentageTarget = lSngPercentageTarget / 100
      lDatStart = Now(): mBolTraining = True
      Call NetSetDestroy 'destroys previous sets
      For i = 0 To mLngPages - 1
        Call PutInputs(i)
        Call PutOutputs(i)
        If NetSetRecord() <> i + 1 Then
          MsgBox "Error in NetSetRecord", vbCritical, App.Title
        End If
      Next i
      Call NetSetStart(0, True)
      For i = 0 To lLngTmp - 1
        Call CaptionUpdate("Learning " & i + 1 & "/" & lLngTmp & " (" & FormatNumber(lSngPercentOk_Prev * 100, 2, True, False, True) & "% success, max.: " & FormatNumber(lSngPercentMax * 100, 2, True, False, True) & "% at cycle " & lLngCycleWhereMaxHappened & ", total time " & DateDiff("s", lDatStart, Now()) & " seconds)")
        'Me.Refresh
        DoEvents
        If mBolTraining Then
          If NetSetLearnStart(PERCENTAGE_OF_ACTIVATION_TO_CONSIDER_ACTIVE, 0, 0) = 0 Then
            For j = 0 To mLngPages - 1
              lSngPercentOk = NetSetLearnContinue(j, True, False)
            Next j
            lSngPercentOk = NetSetLearnEnd()
          Else
            MsgBox "Could not start self training.", vbCritical, App.Title
          End If
          If lSngPercentOk > lSngPercentMax Then
            lSngPercentMax = lSngPercentOk
            Call NetSnapshotTake
            lLngCycleWhereMaxHappened = i
          End If
          If lSngPercentOk >= lSngPercentageTarget Then
            Call CaptionUpdate("Learned (" & FormatNumber(lSngPercentMax * 100, 2, True, False, True) & "% success)")
            Exit For
          End If
          lSngPercentOk_Prev = lSngPercentOk
        Else
          Exit For
        End If
      Next i
      If i > lLngTmp And lSngPercentMax <> 0 Then Call NetSnapshotGet
    End If
  End If
  
End Sub

Private Sub LearnPage()
  Call PutInputs(mLngPage_Current)
  Call PutOutputs(mLngPage_Current)
  Call NetLearn(0)
End Sub

Private Sub PutInputs(ByVal pLngPage As Long)
  Dim i As Long
  Dim j As Long
  Dim lLngCount As Long
  
  For i = 0 To mIntIRows - 1
    For j = 0 To mIntICols - 1
      Call NetInputSet(lLngCount, mPages(pLngPage).SngInputs(i * mIntICols + j))
      lLngCount = lLngCount + 1
    Next j
  Next i

End Sub

Private Sub PutOutputs(ByVal pLngPage As Long)
  Dim i As Long
  Dim j As Long
  Dim lLngCount As Long

  For i = 0 To mIntORows - 1
    For j = 0 To mIntOCols - 1
      Call NetOutputSet(lLngCount, mPages(pLngPage).SngOutputs(i * mIntOCols + j))
      lLngCount = lLngCount + 1
    Next j
  Next i

End Sub

Private Sub CaptionUpdate(ByVal pStrCaption As String)

  If LenB(pStrCaption) <> 0 Then
    Me.Caption = FORM_CAPTION & FORM_CAPTION_SEPARATOR & pStrCaption
  Else
    Me.Caption = FORM_CAPTION
  End If
  
  If mBolThinking Then Me.Caption = Me.Caption & FORM_CAPTION_SEPARATOR & "Thinking"
  
  Me.Caption = Me.Caption & FORM_CAPTION_SEPARATOR & "Type m for help"
  
End Sub

Private Sub MemorizeIOAndCreateNewPaper()
  
  mLngPages = mLngPages + 1
  Call IOReserveMemoryForThisPage
  mLngPage_Current = mLngPages - 1
  Call PaintIOs(picInputs, mIntICols, mIntIRows, mPages(mLngPage_Current).SngInputs)
  Call PaintIOs(picOutputs, mIntOCols, mIntORows, Outputs())
  
End Sub

Private Sub IOClear()
  Dim i As Long
  
  For i = 0 To mIntICols * mIntIRows - 1
    mPages(mLngPage_Current).SngInputs(i) = 0
  Next i
  
  For i = 0 To mIntOCols * mIntORows - 1
    mPages(mLngPage_Current).SngOutputs(i) = 0
  Next i
  
  Call PaintIOs(picInputs, mIntICols, mIntIRows, mPages(mLngPage_Current).SngInputs)
  Call PaintIOs(picOutputs, mIntOCols, mIntORows, Outputs())
  
End Sub

Private Sub IOReserveMemoryForThisPage()

  ReDim Preserve mPages(mLngPages - 1) As tIOPage
  
  With mPages(mLngPages - 1)
    ReDim .SngInputs(mIntICols * mIntIRows - 1) As Single
    ReDim .SngOutputs(mIntOCols * mIntORows - 1) As Single
  End With

End Sub

Private Sub Form_Load()

  'default
  mLngColorGrey = RGB(127, 127, 127)
  mSngInk = 1
  mIntICols = 10
  mIntIRows = 10
  mIntOCols = 1
  mIntORows = 5
  
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mSngPctIO = X / Me.Width
  Call Resize
  Call PaintGrid(picInputs, mIntICols, mIntIRows)
  Call PaintGrid(picOutputs, mIntOCols, mIntORows)
  Call PaintIOs(picInputs, mIntICols, mIntIRows, mPages(mLngPage_Current).SngInputs)
  Call PaintIOs(picOutputs, mIntOCols, mIntORows, Outputs())
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call NetDestroy
End Sub

Private Sub Form_Resize()
  If mBolActivated Then Call UpdateForm
End Sub

Private Sub UpdateForm()
  Call Resize
  Call PaintGrid(picInputs, mIntICols, mIntIRows)
  Call PaintGrid(picOutputs, mIntOCols, mIntORows)
  Call PageRefresh
End Sub

Private Sub Resize()
  Dim lSngTmp As Single
  Dim lLngTmp As Long

  With picInputs
    .Left = 0
    .Top = 0
    .Width = Me.Width * mSngPctIO
    lLngTmp = GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY * 1.2
    .Height = Me.Height - lLngTmp
  End With
  
  With picOutputs
    .Top = 0
    lSngTmp = picInputs.Width + 50
    .Left = lSngTmp
    .Width = Me.Width - lSngTmp
    .Height = Me.Height - lLngTmp
  End With

End Sub

Private Sub PaintGrid(pPic As PictureBox, _
                      ByVal pIntCols As Integer, _
                      ByVal pIntRows As Integer)
  Dim i As Integer
  Dim lSngTmp1 As Single
  Dim lSngTmp2 As Single
    
  pPic.Cls
  
  lSngTmp1 = pPic.Width / pIntCols
  
  For i = 1 To pIntCols - 1
    lSngTmp2 = lSngTmp1 * i
    pPic.Line (lSngTmp2, 0)-(lSngTmp2, pPic.Height), mLngColorGrey
  Next i
  
  lSngTmp1 = pPic.Height / pIntRows
  
  For i = 1 To pIntRows - 1
    lSngTmp2 = lSngTmp1 * i
    pPic.Line (0, lSngTmp2)-(pPic.Width, lSngTmp2), mLngColorGrey
  Next i
    
End Sub

Private Sub PaintIOs(pPic As PictureBox, _
                     ByVal pIntCols As Integer, _
                     ByVal pIntRows As Integer, _
                     pSngIO() As Single)
  Dim i As Integer
  Dim j As Integer
  Dim lSngTmp1 As Single
  Dim lSngTmp2 As Single
  Dim lIntTmp As Integer
    
  With pPic
    lSngTmp1 = .Width / pIntCols
    lSngTmp2 = .Height / pIntRows
  End With
      
  For i = 0 To pIntRows - 1
    For j = 0 To pIntCols - 1
      Call PaintIO(pPic, pIntCols, pIntRows, j, i, pSngIO)
    Next j
  Next i

End Sub

Private Sub PaintIO(pPic As PictureBox, _
                    ByVal pIntCols As Integer, _
                    ByVal pIntRows As Integer, _
                    ByVal pIntCol As Integer, _
                    ByVal pIntRow As Integer, _
                    pSngIO() As Single)
  Dim lSngTmp1 As Single
  Dim lSngTmp2 As Single
  Dim lIntTmp As Integer
  Dim lLngIdx As Long
    
  With pPic
    lSngTmp1 = .Width / pIntCols
    lSngTmp2 = .Height / pIntRows
    lLngIdx = pIntRow * pIntCols + pIntCol
    If CStr(pSngIO(lLngIdx)) <> "1,#QNAN" Then
      lIntTmp = 255 * (1 - Val(Replace$(pSngIO(lLngIdx), ",", ".")))
    Else
      lIntTmp = 255
    End If
    .FillStyle = 0 'solid
    pPic.Line (lSngTmp1 * pIntCol, lSngTmp2 * pIntRow)-(lSngTmp1 * (pIntCol + 1), lSngTmp2 * (pIntRow + 1)), _
              RGB(lIntTmp, lIntTmp, lIntTmp), _
              BF
    .FillStyle = 1 'transparent
    pPic.Line (lSngTmp1 * pIntCol, lSngTmp2 * pIntRow)-(lSngTmp1 * (pIntCol + 1), lSngTmp2 * (pIntRow + 1)), _
              mLngColorGrey, B
  End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub picInputs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call InputsMouseDown(Button, Shift, X, Y)
End Sub

Private Sub InputsMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lIntCol As Integer
  Dim lIntRow As Integer
  Dim lBolTmp As Boolean
  
  If Button = 1 Then
    With picInputs
      lIntCol = Int(X / (.Width / mIntICols))
      lIntRow = Int(Y / (.Height / mIntIRows))
    End With
    lBolTmp = lIntCol >= 0 And lIntRow >= 0 And lIntCol < mIntICols And lIntRow < mIntIRows
  ElseIf Button = 2 Then
    Call ShowHelp
  End If
  If lBolTmp Then
    mPages(mLngPage_Current).SngInputs(lIntRow * mIntICols + lIntCol) = mSngInk
    Call PaintIO(picInputs, mIntICols, mIntIRows, lIntCol, lIntRow, mPages(mLngPage_Current).SngInputs)
    If mBolThinking Then Call PaintIOs(picOutputs, mIntOCols, mIntORows, Outputs())
  End If

End Sub

Private Sub OutputsMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lIntCol As Integer
  Dim lIntRow As Integer
  Dim lBolTmp As Boolean
  Dim lIntTmp As Integer
  
  If Not mBolThinking Then
    If Button = 1 Then
      With picOutputs
        lIntCol = Int(X / (.Width / mIntOCols))
        lIntRow = Int(Y / (.Height / mIntORows))
      End With
      lBolTmp = lIntCol >= 0 And lIntRow >= 0 And lIntCol < mIntOCols And lIntRow < mIntORows
    ElseIf Button = 2 Then
      Call ShowHelp
    End If
    If lBolTmp Then
      mPages(mLngPage_Current).SngOutputs(lIntRow * mIntOCols + lIntCol) = mSngInk
      Call PaintIO(picOutputs, mIntOCols, mIntORows, lIntCol, lIntRow, mPages(mLngPage_Current).SngOutputs)
    End If
  End If
  
End Sub

Private Sub picInputs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call InputsMouseDown(Button, Shift, X, Y)
End Sub

Private Sub picOutputs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call OutputsMouseDown(Button, Shift, X, Y)
End Sub

Private Sub picOutputs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call OutputsMouseDown(Button, Shift, X, Y)
End Sub

