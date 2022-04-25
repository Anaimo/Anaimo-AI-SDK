Imports System.ComponentModel
Imports System.IO

Public Class Form1
  Const TIME_TO_UPDATE_PROGRESS = 2
  Const MODES_STR = "0 For Standard back propagation, 1 For Optimized BackProp, 2 For Dynamic propagation, "

  Private mSngPctIO As Single

  Private mIntICols As Integer = 10 'number of input columns
  Private mIntIRows As Integer = 10 'number of input rows

  Private mIntOCols As Integer = 1 'number of output columns
  Private mIntORows As Integer = 5 'number of output rows

  Private mPages() As tIOPage
  Private mLngPages As Long 'total number of pages
  Private mLngPage_Current As Long 'current page

  Private mBolWorking As Boolean
  Private mBolActivated As Boolean

  Private mBolThinking As Boolean 'showing outputs (false) or thinking the outputs (true)

  Private Const MAX_NUMBER_OF_NEURONS = 2000000

  Private Const PERCENTAGE_OF_ACTIVATION_TO_CONSIDER_ACTIVE = 0.5

  Private mLngColorGrey As Long = 127

  Private Const MAX_COLOR = 16777215

  Private mLngColor As Single = 0

  Private mBolBlackAndWhite As Boolean = True

  Private mLngActivationFunction As Long

  'current neural network configuration
  Private mIntLayers As Integer

  Private Const ActivationFunction_Sigmoid = 0
  Private Const ActivationFunction_ReLU = 1
  Private Const ActivationFunction_FastSigmoid = 2
  Private Const ActivationFunction_END = 3

  Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
    Dim lBolUpdate As Boolean
    Dim lStrTmp As String
    Dim lSngTmp As Single
    Dim lLngTmp As Long

    Cursor.Current = Cursors.WaitCursor

    If e.Modifiers = Keys.Control Then
      If mLngColor < MAX_COLOR * PERCENTAGE_OF_ACTIVATION_TO_CONSIDER_ACTIVE Then mLngColor = MAX_COLOR Else mLngColor = 0
    ElseIf e.KeyCode = Keys.B Then
      mBolBlackAndWhite = Not mBolBlackAndWhite
      lBolUpdate = True
    ElseIf e.KeyCode = Keys.F Then
      mLngActivationFunction = (mLngActivationFunction + 1) Mod ActivationFunction_END
      Call NetActivationFunctionSet(mLngActivationFunction)
      lStrTmp = "Activation Function Is now: "
      Select Case mLngActivationFunction
        Case ActivationFunction_Sigmoid : lStrTmp &= "Sigmoid"
        Case ActivationFunction_ReLU : lStrTmp &= "ReLU"
        Case ActivationFunction_FastSigmoid : lStrTmp &= "FastSigmoid"
      End Select
      MsgBox(lStrTmp, vbInformation)
      lBolUpdate = True
    ElseIf e.KeyCode = Keys.Escape Then
      'stop studying
      mBolWorking = False
      lBolUpdate = True
    ElseIf e.KeyCode = Keys.K Then
      Call InkSelect()
    ElseIf e.KeyCode = Keys.I Then 'initialize
      Call Initialize()
    ElseIf e.KeyCode = Keys.Y Then 'initialize number of layers
      Call InitializeWithANewNumberOfLayers(mIntICols, mIntIRows, mIntOCols, mIntORows, False)
    ElseIf e.KeyCode = Keys.C Then 'clear
      Call IOClear()
    ElseIf e.KeyCode = Keys.N Then 'new
      Call MemorizeIOAndCreateNewPage()
      Call PaintPage()
      Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, (mLngPage_Current + 1) & "/" & mLngPages, False, False)
    ElseIf e.KeyCode = Keys.Home Then 'go to 1st page
      mLngPage_Current = 0
      lBolUpdate = True
    ElseIf e.KeyCode = Keys.End Then 'key to go to last page
      mLngPage_Current = mLngPages - 1
      lBolUpdate = True
    ElseIf e.KeyCode = Keys.Right Then 'right arrow
      If mLngPage_Current < mLngPages - 1 Then
        mLngPage_Current += 1
        lBolUpdate = True
      End If
    ElseIf e.KeyCode = Keys.Left Then 'left arrow
      If mLngPage_Current > 0 Then
        mLngPage_Current -= 1
        lBolUpdate = True
      End If
    ElseIf e.KeyCode = Keys.A Then 'learn set
      Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, (mLngPage_Current + 1) & "/" & mLngPages & FORM_CAPTION_SEPARATOR & " Studying", False, False)
      Call PagesLearn()
      Call PaintIOs(picInputs, mIntICols, mIntIRows, mPages(mLngPage_Current).SngInputs, False)
      Call PaintIOs(picOutputs, mIntOCols, mIntORows, Outputs(False), True)
    ElseIf e.KeyCode = Keys.P Then 'learn page
      Call LearnPage()
      Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, (mLngPage_Current + 1) & "/" & mLngPages & FORM_CAPTION_SEPARATOR & " Learned", False, False)
    ElseIf e.KeyCode = Keys.T Then 'think and say what are these inputs
      mBolThinking = Not mBolThinking
      lBolUpdate = True
    ElseIf e.KeyCode = Keys.D Then 'delete current page
      If mLngPages > 1 Then
        Call PageDelete()
        lBolUpdate = True
      End If
    ElseIf e.KeyCode = Keys.M Then
      Call ShowHelp()
    ElseIf e.KeyCode = Keys.H Then
      MsgBox("Hardware id of this device is (you can paste it as has been copied to your clipboard): " & MyHardwareIdAndIntoClipBoard(), vbInformation)
    ElseIf e.KeyCode = Keys.R Then
      Call PagesCreateRandom()
    ElseIf e.KeyCode = Keys.V Then 'verify matching of inputs & outputs
      Call PagesVerify()
    ElseIf e.KeyCode = Keys.S Then 'save all pages
      If SaveFile() Then Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, (mLngPage_Current + 1) & "/" & mLngPages & FORM_CAPTION_SEPARATOR & " Saved to file", False, False)
    ElseIf e.KeyCode = Keys.O Then
      If LoadFile() Then Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, "Loaded from file", False, False)
    ElseIf e.KeyCode = Keys.J Then
      lBolUpdate = LoadFolder()
    ElseIf e.KeyCode = Keys.L Then
      If ConfigureParameter("Learning rate", 0, 1, NetLearningRateGet(), lSngTmp) Then
        NetLearningRateSet(lSngTmp)
      End If
    ElseIf e.KeyCode = Keys.U Then
      If ConfigureParameter("Momentum", 0, 1, NetMomentumGet(), lSngTmp) Then
        NetMomentumSet(lSngTmp)
      End If
    ElseIf e.KeyCode = Keys.X Then
      If ConfigureParameter("Maximum number of threads", 1, 4096, NetThreadsMaxNumberGet(), lSngTmp) Then
        NetThreadsMaxNumberSet(CInt(lSngTmp))
      End If
    ElseIf e.KeyCode = Keys.W Then
      If ConfigureParameter("Mode (" & MODES_STR & ")", MODE_STANDARD_BACKPROPAGATION, MODE_FINISH, NetModeGet(), lSngTmp) Then
        NetModeSet(CInt(lSngTmp))
      End If
    ElseIf e.KeyCode = Keys.E Then
      MsgBox("Current number of layers is: " & mIntLayers, vbInformation, "Neural network layers")
    ElseIf e.KeyCode = Keys.G Then
      lLngTmp = Val(InputBox("Move to page?"))
      If lLngTmp >= 1 And lLngTmp <= mLngPages Then
        mLngPage_Current = lLngTmp - 1
        lBolUpdate = True
      End If
    End If
    If lBolUpdate Then Call PageRefresh(True)

    Cursor.Current = Cursors.Default

  End Sub

  Private Function ConfigureParameter(ByVal pStrName As String,
                                      ByVal pSngMin As Single,
                                      ByVal pSngMax As Single,
                                      ByVal pSngParameterOriginal As Single,
                                      ByRef pSngParameterEntered As Single) As Boolean
    Dim lStrTmp As String
    Dim lSngTmp As Single

    lStrTmp = InputBox(pStrName & "? [" & pSngMin & ", " & pSngMax & "]", pStrName, pSngParameterOriginal)
    If IsNumeric(lStrTmp) Then
      lSngTmp = CSng(lStrTmp)
      If lSngTmp > pSngMax Then
        pSngParameterEntered = pSngMax
      ElseIf lSngTmp < pSngMin Then
        pSngParameterEntered = pSngMin
      Else
        pSngParameterEntered = lSngTmp
      End If
      Return True
    Else
      Return False
    End If

  End Function

  Private Sub ShowHelp()

    MsgBox("Mouse click to draw, or type: " & vbCrLf & vbCrLf &
     "Left & Right arrows, Start, End or g: moves between pages." & vbCrLf &
     "CTRL or SHIFT: switches from drawing or erasing." & vbCrLf &
     "b: switches black & white (B&W) or color mode." & vbCrLf &
     "k: changes the ink color." & vbCrLf &
     "i: initializes neural network or everything." & vbCrLf &
     "y: initializes with a different number of layers." & vbCrLf &
     "c: clears current page." & vbCrLf &
     "n: creates new page." & vbCrLf &
     "r: creates random pages." & vbCrLf &
     "d: deletes current page." & vbCrLf &
     "f: changes the learning function." & vbCrLf &
     "l: informs/sets the learning rate." & vbCrLf &
     "u: informs/sets the momentum." & vbCrLf &
     "x: informs/sets the maximum number of threads." & vbCrLf &
     "w: informs/sets the mode (" & MODES_STR & ")." & vbCrLf &
     "e: informs of the current number of layers." & vbCrLf &
     "p: learns current page." & vbCrLf &
     "a: learns all pages." & vbCrLf &
     "ESC: stops." & vbCrLf &
     "t: thinks to show outputs." & vbCrLf &
     "v: verifies pages versus thought." & vbCrLf &
     "s: saves pages or knowledge." & vbCrLf &
     "o: loads pages or knowledge." & vbCrLf &
     "j: opens all the jpg files in a folder and adds them to the current pages." & vbCrLf &
     "h: shows hardware id of this device and you can paste it." & vbCrLf &
     "m: shows this help.", vbInformation, My.Application.Info.Title)

    Call UpdateForm(False)

  End Sub

  Private Function InitializeWithANewNumberOfLayers(ByVal pIntICols As Integer,
                                                    ByVal pIntIRows As Integer,
                                                    ByVal pIntOCols As Integer,
                                                    ByVal pIntORows As Integer,
                                                    ByVal pBolSilent As Boolean) As Boolean
    Dim lIntLayers As Integer

    If pBolSilent Then
      If mIntLayers = 0 Then lIntLayers = pIntORows + 1 Else lIntLayers = mIntLayers
    Else
      lIntLayers = Val(InputBox("Number of layers?", "Number of layers", IIf(mIntLayers = 0, pIntORows + 1, mIntLayers)))
    End If

    If lIntLayers >= 2 Then
      If NeuralNetworkCreate(pIntICols, pIntIRows, pIntOCols, pIntORows, lIntLayers) Then
        mIntLayers = lIntLayers
        Return True
      Else
        Return False
      End If
    Else
      Return False
    End If

  End Function

  Private Function InitializeAutomatically() As Boolean

    If InitializeWithANewNumberOfLayers(mIntICols, mIntIRows, mIntOCols, mIntORows, True) Then
      Call InitializeInternals()
      Return True
    Else
      Return False
    End If

  End Function

  Private Sub Initialize()
    Dim lIntICols As Integer
    Dim lIntIRows As Integer
    Dim lIntOCols As Integer
    Dim lIntORows As Integer
    Dim lIntRes As Integer

    lIntRes = MsgBox("Initialize also the number of inputs, outputs and layers?", vbQuestion + vbYesNoCancel)
    If lIntRes = vbYes Then
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
                Call InitializeInternals()
              End If
            End If
          End If
        End If
      End If
    ElseIf lIntRes = vbNo Then
      Call NetInitialize(0)
    End If

  End Sub

  Private Sub InitializeInternals()

    mLngPages = 1 'first page
    mLngPage_Current = 0
    ReDim mPages(mLngPages - 1) 'adjusts memory
    mSngPctIO = picInputs.Width / Me.Width
    picInputs.Visible = True
    picOutputs.Visible = True
    Call IOReserveMemoryForThisPage()
    Call UpdateForm(True)
    Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, (mLngPage_Current + 1) & "/" & mLngPages, False, False)

  End Sub

  Private Function SaveFile() As Boolean
    Dim lBolReturn As Boolean

    With SaveFileDialog1
      .Filter = "NNViewer Project (*.nnv)|*.nnv|NNViewer Knowledge (*.nnk)|*.nnk"
      .InitialDirectory = Application.StartupPath()
      .FileName = vbNullString
      .ShowDialog()
      If Len(.FileName) <> 0 Then
        Cursor.Current = Cursors.WaitCursor
        Select Case Strings.Right(.FileName, 3)
          Case "nnv" : lBolReturn = PagesSaveToFile(.FileName)
          Case "nnk" : lBolReturn = KnowSaveToFile(.FileName)
        End Select
        Cursor.Current = Cursors.Default
      End If
    End With

    Return lBolReturn

  End Function

  Private Function LoadFile() As Boolean
    Dim lBolReturn As Boolean

    With OpenFileDialog1
      .Filter = "NNViewer Project (*.nnv)|*.nnv|NNViewer Knowledge (*.nnk)|*.nnk|Icon (*.ico)|*.ico"
      .InitialDirectory = Application.StartupPath()
      .FileName = vbNullString
      .ShowDialog()
      If Len(.FileName) <> 0 Then
        Cursor.Current = Cursors.WaitCursor
        Select Case LCase$(Strings.Right$(.FileName, 4))
          Case ".nnv" : lBolReturn = LoadFile_NNV(.FileName)
          Case ".ico" : lBolReturn = LoadFile_ICO(.FileName)
          Case ".nnk" : lBolReturn = LoadFile_NNK(.FileName)
        End Select
        Cursor.Current = Cursors.Default
      End If
    End With

    Return lBolReturn

  End Function

  Private Function LoadFolder() As Boolean
    Dim lBolReturn As Boolean
    Dim lLngStepMax As Long
    Dim lLngStep As Long
    Dim lIntWidthMax As Integer
    Dim lIntHeightMax As Integer
    Dim lIntResponse As Integer
    Dim lIntActivateOutput As Integer
    Dim lImage As Image

    With FolderBrowserDialog1
      .ShowNewFolderButton = False ' Disable the creation of new folders

      ' Open the folder we want
      If .ShowDialog = Windows.Forms.DialogResult.OK Then

        lLngStepMax = My.Computer.FileSystem.GetFiles(.SelectedPath,
                FileIO.SearchOption.SearchTopLevelOnly,
                "*.jpg").Count

        If lLngStepMax <= 0 Then
          MsgBox("Folder does not seem to have any jpg image.", vbExclamation)
        Else
          ' Access to all the images
          For Each lStrImg As String In My.Computer.FileSystem.GetFiles(.SelectedPath,
                FileIO.SearchOption.SearchTopLevelOnly,
                "*.jpg")
            lImage = Image.FromFile(lStrImg)
            ' Take greater height and width
            If lImage.Size.Width > lIntWidthMax Then lIntWidthMax = lImage.Size.Width
            If lImage.Size.Height > lIntHeightMax Then lIntHeightMax = lImage.Size.Height
            lLngStep += 1
            If lLngStep Mod 10 = 0 Then
              GC.Collect() 'free unused memory
              Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, "Analyzing image " & (lLngStep + 1) & "/" & lLngStepMax & "...", True, False)
            End If
            lImage = Nothing
          Next

          lLngStep = 0

          ' If height or width are greater than current NN inputs we create a new grid
          If lIntWidthMax > mIntICols Or lIntHeightMax > mIntIRows Then
            lIntResponse = MsgBox("Maximum height (" & lIntHeightMax & ") or width (" & lIntWidthMax & ") of the images are greater than current rows (" & mIntIRows & ") or columns (" & mIntICols & ") of the neural network. Do you want to continue by rescaling the images? (answering No cancels)", MsgBoxStyle.YesNo + MessageBoxIcon.Question, "Rescale images")
          Else
            lIntResponse = MsgBoxResult.Yes
          End If
          If lIntResponse = MsgBoxResult.Yes Then
            lIntActivateOutput = Val(InputBox("Output to activate? (0 for none)", "Activate output", 1))
            If lIntActivateOutput >= 0 Then
              With picTmp
                .Width = mIntICols
                .Height = mIntIRows
              End With
              For Each lStrImg As String In My.Computer.FileSystem.GetFiles(.SelectedPath,
                FileIO.SearchOption.SearchTopLevelOnly,
                "*.jpg")
                Call MemorizeIOAndCreateNewPage()
                Call ImageLoadIntoPicTmp(lStrImg, True)
                Call ImageLoadIntoPage()
                If lIntActivateOutput > 0 Then
                  mPages(mLngPage_Current).SngOutputs(lIntActivateOutput - 1) = 1
                End If
                If lLngStep Mod 10 = 0 Then
                  GC.Collect() 'free unused memory
                  Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, "Storing image " & (lLngStep + 1) & "/" & lLngStepMax & " (total pages " & mLngPage_Current & ")...", True, False)
                End If
                lLngStep += 1
              Next
              Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, (mLngPage_Current + 1) & "/" & mLngPages, False, False)
              Call PageRefresh(True)
            End If
          End If
        End If
      End If
    End With

    Return lBolReturn

  End Function

  Private Function ResizeImage(ByVal InputImage As Image) As Image

    Return New Bitmap(InputImage, New Drawing.Size(picTmp.Width, picTmp.Height))
  End Function

  Private Function LoadFile_NNV(ByVal pStrFileName As String) As Boolean
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
    Dim lFileSource As FileStream
    Dim lFileReader As StreamReader
    Dim lBolReturn As Boolean

    'default
    lBolError = True

    If System.IO.File.Exists(pStrFileName) Then
      lFileSource = New FileStream(pStrFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
      lFileReader = New StreamReader(lFileSource)
      lStrTmp = lFileReader.ReadLine()
      lPieces = Split(lStrTmp, ";")
      If SafeUbound(lPieces) = 4 Then
        lIntICols = lPieces(0)
        lIntIRows = lPieces(1)
        lIntOCols = lPieces(2)
        lIntORows = lPieces(3)
        If InitializeWithANewNumberOfLayers(lIntICols, lIntIRows, lIntOCols, lIntORows, False) Then
          lLngPages = lPieces(4)
          ReDim lPages(lLngPages - 1)
          lBolError = False
          For p = 0 To lLngPages - 1
            With lPages(p)
              ReDim .SngInputs(lIntICols * lIntIRows - 1)
              ReDim .SngOutputs(lIntOCols * lIntORows - 1)
            End With
            For i = 0 To lIntIRows - 1
              lStrTmp = lFileReader.ReadLine()
              lPieces = Split(lStrTmp, ";")
              If SafeUbound(lPieces) >= lIntICols Then
                For j = 0 To lIntICols - 1
                  lPages(p).SngInputs(i * lIntICols + j) = lPieces(j)
                Next j
              Else
                MsgBox("Incorrect format in page " & p & " input row " & i, vbCritical)
                lBolError = True
                Exit For
              End If
            Next i
            If Not lBolError Then
              For i = 0 To lIntORows - 1
                lStrTmp = lFileReader.ReadLine()
                lPieces = Split(lStrTmp, ";")
                If SafeUbound(lPieces) >= lIntOCols Then
                  For j = 0 To lIntOCols - 1
                    lPages(p).SngOutputs(i * lIntOCols + j) = lPieces(j)
                  Next j
                Else
                  MsgBox("Incorrect format in page " & p & " output row " & i, vbCritical)
                  lBolError = True
                  Exit For
                End If
              Next i
            End If
            If lBolError Then Exit For
          Next p
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
            lBolReturn = True 'success
          End If
        End If
      Else
        MsgBox("Incorrect format: header should contain the number of input columns, input rows, output columns and output rows.", vbCritical)
      End If
      lFileReader.Close()
      lFileSource.Close()
    Else
      MsgBox("File Does Not Exist")
    End If

    Return lBolReturn

  End Function

  Private Sub ImageLoadIntoPicTmp(ByVal pStrFileName As String,
                                  ByVal pBolImageResize As Boolean)
    Dim lImage As Image

    If pBolImageResize Then
      picTmp.SizeMode = PictureBoxSizeMode.StretchImage
    Else
      picTmp.SizeMode = PictureBoxSizeMode.Normal
    End If

    lImage = Image.FromFile(pStrFileName)

    picTmp.Image = lImage

  End Sub

  Private Function LoadFile_ICO(ByVal pStrFileName As String) As Boolean
    Dim lIntCols As Integer
    Dim lIntRows As Integer
    Dim i As Long
    Dim lBolReturn As Boolean

    'clears inputs
    For i = 0 To mIntICols * mIntIRows - 1
      mPages(mLngPage_Current).SngInputs(i) = 0
    Next i

    ImageLoadIntoPicTmp(pStrFileName, False)

    'lIntCols = picTmp.Width - 2
    'lIntRows = picTmp.Height - 2

    If lIntCols <> mIntICols Or lIntRows <> mIntIRows Then
      If MsgBox("Current inputs grid rows (" & mIntIRows & ") or columns (" & mIntICols & ") do not match the icon's rows (" & lIntRows & ") or columns (" & lIntCols & "). Would you like to adjust grid and neural network?", vbQuestion + vbYesNo) = vbYes Then
        If InitializeWithANewNumberOfLayers(lIntCols, lIntRows, mIntOCols, mIntORows, False) Then
          mIntICols = lIntCols
          mIntIRows = lIntRows
          Call InitializeInternals()
          lBolReturn = True
        End If
      End If
    Else
      lBolReturn = True
    End If

    If lBolReturn Then
      Call ImageLoadIntoPage()
      Call PaintIOs(picInputs, mIntICols, mIntIRows, mPages(mLngPage_Current).SngInputs, False)
      picTmp.Visible = False
    End If

    Return lBolReturn

  End Function

  Private Sub ImageLoadIntoPage()
    Dim i, j As Long
    Dim lBitmap As Bitmap
    Dim lColor As Color

    lBitmap = CType(picTmp.Image, Bitmap)

    lBitmap = ResizeImage(lBitmap)

    For i = 0 To picTmp.ClientRectangle.Height - 1
      For j = 0 To picTmp.ClientRectangle.Width - 1
        lColor = lBitmap.GetPixel(j, i)
        mPages(mLngPage_Current).SngInputs(i * mIntICols + j) = 1 - RGB(lColor.R, lColor.G, lColor.B) / MAX_COLOR
      Next j
    Next i

  End Sub

  Private Function PagesSaveToFile(ByVal pStrFileName As String) As Boolean
    Dim i, j, p As Long
    Dim lStrTmp As String
    Dim lFile As System.IO.StreamWriter

    lFile = My.Computer.FileSystem.OpenTextFileWriter(pStrFileName, False)
    If (Not lFile Is Nothing) Then
      lFile.WriteLine(mIntICols & ";" & mIntIRows & ";" & mIntOCols & ";" & mIntORows & ";" & mLngPages)
      For p = 0 To mLngPages - 1
        For i = 0 To mIntIRows - 1
          lStrTmp = vbNullString
          For j = 0 To mIntICols - 1
            lStrTmp = lStrTmp & mPages(p).SngInputs(i * mIntICols + j) & ";"
          Next j
          lFile.WriteLine(lStrTmp)
        Next i
        For i = 0 To mIntORows - 1
          lStrTmp = vbNullString
          For j = 0 To mIntOCols - 1
            lStrTmp = lStrTmp & mPages(p).SngOutputs(i * mIntOCols + j) & ";"
          Next j
          lFile.WriteLine(lStrTmp)
        Next i
      Next p
      lFile.Close()
      Return True
    End If

  End Function

  Private Function LoadFile_NNK(ByVal pStrFileName As String) As Boolean
    Dim lStrTmp As String
    Dim lFileSource As FileStream
    Dim lFileReader As StreamReader
    Dim lBolReturn As Boolean
    Dim lPieces() As String
    Dim lLngNeuronsMaxNumber, lLngNetInputsMaxNumber, lLngNetOutputsMaxNumberGet As Long
    Dim lLngTotalNumberOfRecords As Long
    Dim lLngRecord As Long
    Dim lDatTmp, lDatNow As Date
    Dim lLngNeuron As Long
    Dim lLngNumberOfInputs As Long
    Dim lLngInput As Long

    If System.IO.File.Exists(pStrFileName) Then
      'default: ok
      lBolReturn = True
      lFileSource = New FileStream(pStrFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
      lFileReader = New StreamReader(lFileSource)
      lStrTmp = lFileReader.ReadLine()
      lPieces = Split(lStrTmp, ";")
      If lPieces.Length <> 5 Then
        MsgBox("File has an incorrect format in line 1.", vbCritical)
        lBolReturn = False
      Else
        lLngTotalNumberOfRecords = lPieces(0)
        lLngNeuronsMaxNumber = NetNeuronsMaxNumberGet()
        lLngNetInputsMaxNumber = NetInputsMaxNumberGet()
        lLngNetOutputsMaxNumberGet = NetOutputsMaxNumberGet()
        If lPieces(1) <> lLngNeuronsMaxNumber Then
          MsgBox("The maximum number of neurons of the file (" & lPieces(1) & ") does not match with current (" & lLngNeuronsMaxNumber & ")", vbExclamation)
          lBolReturn = False
        ElseIf lPieces(2) <> lLngNetInputsMaxNumber Then
          MsgBox("The maximum number of inputs of the file (" & lPieces(2) & ") does not match with current (" & lLngNetInputsMaxNumber & ")", vbExclamation)
          lBolReturn = False
        ElseIf lPieces(3) <> lLngNetOutputsMaxNumberGet Then
          MsgBox("The maximum number of outputs of the file (" & lPieces(3) & ") does not match with current (" & lLngNetOutputsMaxNumberGet & ")", vbExclamation)
          lBolReturn = False
        ElseIf lPieces(4) <> mIntLayers Then
          MsgBox("File was created from " & lPieces(4) & " layers and currently there are " & mIntLayers & ".", vbExclamation)
          lBolReturn = False
        ElseIf lFileReader.EndOfStream Then
          MsgBox("File has no data", vbCritical)
        Else
          While Not lFileReader.EndOfStream And lBolReturn
            lStrTmp = lFileReader.ReadLine()
            lLngRecord += 1
            lPieces = Split(lStrTmp, ";")
            If lPieces.Length <> 3 Then
              MsgBox("File has an incorrect format in record " & lLngRecord & ".", vbCritical)
              lBolReturn = False
            Else
              lLngNeuron = lPieces(0)
              NeuBiasSet(lLngNeuron, lPieces(1))
              lLngNumberOfInputs = CLng(lPieces(2))
              'if it has inputs
              If lLngNumberOfInputs <> 0 Then
                lLngInput = 0
                While Not lFileReader.EndOfStream And lBolReturn And lLngInput < lLngNumberOfInputs
                  lStrTmp = lFileReader.ReadLine()
                  lPieces = Split(lStrTmp, ";")
                  If lPieces.Length <> 3 Then
                    MsgBox("File has an incorrect format in record " & lLngRecord & ".", vbCritical)
                    lBolReturn = False
                  ElseIf lPieces(0) <> lLngNeuron Then
                    MsgBox("Weight in record " & lLngRecord & " of neuron " & lLngNeuron & " was written for neuron " & lPieces(0) & ".", vbCritical)
                    lBolReturn = False
                  Else
                    NeuInputWeightSet(lLngNeuron, lPieces(1), CSng(lPieces(2)))
                  End If
                  lDatNow = Now()
                  If DateDiff(DateInterval.Second, lDatTmp, lDatNow) > TIME_TO_UPDATE_PROGRESS Then
                    lDatTmp = lDatNow
                    Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, "Loading knowledge (" & FormatNumber(lLngRecord / lLngTotalNumberOfRecords * 100, 2, True, False, True) & "%)... ", True, False)
                  End If
                  lLngRecord += 1
                  lLngInput += 1
                End While
              End If
            End If
          End While
        End If
      End If
      lFileReader.Close()
      lFileSource.Close()
    Else
      MsgBox("File Does Not Exist")
      lBolReturn = False
    End If

    Return lBolReturn

  End Function

  Private Function KnowSaveToFile(ByVal pStrFileName As String) As Boolean
    Dim i, j As Long
    Dim lFile As System.IO.StreamWriter
    Dim lLngNeuronsMaxNumber, lLngNetInputsMaxNumber, lLngNetOutputsMaxNumberGet As Long
    Dim lLngTotalNumberOfRecords As Long
    Dim lLngRecord As Long
    Dim lDatTmp, lDatNow As Date
    Dim lLngTotalNumberOfValidWeights As Long
    Dim lLngTmp As Long

    lLngNeuronsMaxNumber = NetNeuronsMaxNumberGet()
    lLngNetInputsMaxNumber = NetInputsMaxNumberGet()
    lLngNetOutputsMaxNumberGet = NetOutputsMaxNumberGet()

    lLngTotalNumberOfRecords = 1

    'calculates total steps and records in database
    For i = 0 To lLngNeuronsMaxNumber - 1
      If (NeuValueUpdatedGet(i) > 0) Then
        lLngTotalNumberOfRecords += 1
        For j = 0 To NeuInputsNumberGet(i) - 1
          If NeuInputWeightUpdatedGet(i, j) > 0 Then
            lLngTotalNumberOfRecords += 1
          End If
        Next j
      End If
    Next i

    lFile = My.Computer.FileSystem.OpenTextFileWriter(pStrFileName, False)

    If (Not lFile Is Nothing) Then
      lFile.WriteLine(lLngTotalNumberOfRecords & ";" & lLngNeuronsMaxNumber & ";" & lLngNetInputsMaxNumber & ";" & lLngNetOutputsMaxNumberGet & ";" & mIntLayers)
      For i = 0 To lLngNeuronsMaxNumber - 1
        If (NeuValueUpdatedGet(i) > 0) Then
          lLngTmp = NeuInputsNumberGet(i)
          If lLngTmp <> 0 Then
            lLngTotalNumberOfValidWeights = 0
            'calculates the number of inputs which will be written
            For j = 0 To NeuInputsNumberGet(i) - 1
              If NeuInputWeightUpdatedGet(i, j) > 0 Then lLngTotalNumberOfValidWeights += 1
            Next j
          End If
          lFile.WriteLine(i & ";" & NeuBiasGet(i) & ";" & lLngTotalNumberOfValidWeights)
          lLngRecord += 1
          For j = 0 To NeuInputsNumberGet(i) - 1
            If NeuInputWeightUpdatedGet(i, j) > 0 Then
              lFile.WriteLine(i & ";" & j & ";" & NeuInputWeightGet(i, j))
              lLngRecord += 1
            End If
          Next j
          lDatNow = Now()
          If DateDiff(DateInterval.Second, lDatTmp, lDatNow) > TIME_TO_UPDATE_PROGRESS Then
            lDatTmp = lDatNow
            Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, "Saving knowledge (" & FormatNumber(lLngRecord / lLngTotalNumberOfRecords * 100, 2, True, False, True) & "%)...", True, False)
          End If
        End If
      Next i
      lFile.Close()
      Return True
    End If

  End Function

  Private Function NeuralNetworkCreate(ByVal pIntICols As Integer,
                                       ByVal pIntIRows As Integer,
                                       ByVal pIntOCols As Integer,
                                       ByVal pIntORows As Integer,
                                       ByVal pIntNumberOfLayers As Integer) As Boolean
    Dim lStrTmp As String = vbNullString

    Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, "Creating neural network...", True, False)

    If CreateNeuralNetwork(Me, MAX_NUMBER_OF_NEURONS, pIntICols * pIntIRows, pIntOCols * pIntORows, 0, pIntNumberOfLayers, pIntNumberOfLayers, lStrTmp) Then
      Call NetInitialize(0)
      MsgBox("Neural network was created successfully with nodes: " & vbCrLf & vbCrLf & lStrTmp, vbInformation, My.Application.Info.Title)
      Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, (mLngPage_Current + 1) & "/" & mLngPages, True, False)
      Return True 'success
    Else
      MsgBox("Error creating neural network.", vbCritical)
      Return False
    End If

  End Function

  Private Sub InkSelect()
    Dim lStrR As String = InputBox("Red value (0 to 255)?", "Red value", 0)
    Dim lStrG As String = InputBox("Green value (0 to 255)?", "Green value", 0)
    Dim lStrB As String = InputBox("Blue value (0 to 255)?", "Blue value", 0)

    If Len(lStrR) <> 0 And Len(lStrG) <> 0 And Len(lStrB) <> 0 Then
      If IsNumeric(lStrR) And IsNumeric(lStrG) And IsNumeric(lStrB) Then
        If Val(lStrR) > 255 Then lStrR = "255"
        If Val(lStrG) > 255 Then lStrG = "255"
        If Val(lStrB) > 255 Then lStrB = "255"
        If Val(lStrR) < 0 Then lStrR = "0"
        If Val(lStrG) < 0 Then lStrG = "0"
        If Val(lStrB) < 0 Then lStrB = "0"
      End If
      mLngColor = RGB(Val(lStrR), Val(lStrG), Val(lStrB))
    End If

  End Sub

  Private Sub PageRefresh(ByVal pBolCaptionRefresh As Boolean)
    Call PaintIOs(picInputs, mIntICols, mIntIRows, mPages(mLngPage_Current).SngInputs, False)
    Call PaintIOs(picOutputs, mIntOCols, mIntORows, Outputs(False), True)
    If pBolCaptionRefresh Then Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, (mLngPage_Current + 1) & "/" & mLngPages, False, False)
  End Sub

  Private Sub PagesVerifyBasedOnHorizontalVisualDivision()
    Dim i, j, k As Long
    Dim lLngTotalGood As Long
    Dim lLngPage_Current_Prev As Long
    Dim lBolThisPageIsCorrect As Boolean
    Dim lSngOutputs() As Single
    Dim lLngTmp As Long
    Dim lSngValueToConsiderActive As Single
    Dim lStrTmp As String = InputBox("Value to consider active?", "Value for active", PERCENTAGE_OF_ACTIVATION_TO_CONSIDER_ACTIVE)

    If Len(lStrTmp) <> 0 Then
      lSngValueToConsiderActive = lStrTmp
      If lSngValueToConsiderActive <> 0 Then
        lStrTmp = vbNullString
        lLngPage_Current_Prev = mLngPage_Current
        lLngTmp = (mIntICols * mIntIRows / mIntORows)
        For i = 0 To mLngPages - 1
          mLngPage_Current = i
          lSngOutputs = Outputs(True)
          For j = 0 To mIntORows - 1
            'if output is not activated, then all corresponding inputs must also be not activated
            If lSngOutputs(j) < lSngValueToConsiderActive Then
              'default
              lBolThisPageIsCorrect = True
              For k = j * lLngTmp To (j + 1) * lLngTmp - 1
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
              For k = j * lLngTmp To (j + 1) * lLngTmp - 1
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
        If Len(lStrTmp) <> 0 Then lStrTmp = ", errors in pages: " & Strings.Left$(lStrTmp, Len(lStrTmp) - 1)
        MsgBox("Percentage of matches is " & FormatNumber(lLngTotalGood / mLngPages * 100, 2, True, False, False) & "%" & lStrTmp, vbInformation, My.Application.Info.Title)
      End If
    End If

  End Sub

  Private Sub PagesVerify()
    Dim i, j As Long
    Dim lLngTotalGood As Long
    Dim lLngPage_Current_Prev As Long
    Dim lBolThisPageIsCorrect As Boolean
    Dim lSngOutputs() As Single
    Dim lLngTmp As Long
    Dim lSngValueToConsiderActive As Single
    Dim lStrTmp As String = InputBox("Value to consider active?", "Value for active", PERCENTAGE_OF_ACTIVATION_TO_CONSIDER_ACTIVE)

    mBolWorking = False 'stops previous

    If Len(lStrTmp) <> 0 Then
      lSngValueToConsiderActive = lStrTmp
      If lSngValueToConsiderActive <> 0 Then
        mBolWorking = True
        lStrTmp = vbNullString
        lLngPage_Current_Prev = mLngPage_Current
        lLngTmp = (mIntICols * mIntIRows / mIntORows)
        For i = 0 To mLngPages - 1
          mLngPage_Current = i
          lSngOutputs = Outputs(True)
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
          If lBolThisPageIsCorrect Then lLngTotalGood += 1 Else lStrTmp = lStrTmp & i + 1 & ","
          Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, "Validating page " & i + 1 & "/" & mLngPages & "...", True, True)
          If Not mBolWorking Then Exit For
        Next i
        mLngPage_Current = lLngPage_Current_Prev
        If Len(lStrTmp) <> 0 Then lStrTmp = ", errors in pages: " & Strings.Left$(lStrTmp, Len(lStrTmp) - 1)
        MsgBox("Percentage of matches is " & FormatNumber(lLngTotalGood / mLngPages * 100, 2, True, False, False) & "%" & lStrTmp, vbInformation, My.Application.Info.Title)
        mBolWorking = False
      End If
    End If

  End Sub

  Private Sub PagesCreateRandom()
    Dim lLngTmp1 As Long
    Dim lLngTmp2 As Long
    Dim lSngTmp1 As Single
    Dim lLngTmp3 As Long
    Dim i, j As Long

    If mIntOCols <> 1 Then
      MsgBox("Random pages can only be created when there is only 1 column of outputs.", vbCritical)
    Else
      lLngTmp1 = Val(InputBox("How many random pages do you want to create?", "Number of random pages", 100))
      If lLngTmp1 > 0 Then
        lSngTmp1 = Val(InputBox("Error rate in percentage?", "Error rate", 0))
        Cursor.Current = Cursors.WaitCursor
        Randomize()
        For i = 1 To lLngTmp1
          mLngPages += 1
          Call IOReserveMemoryForThisPage()
          mLngPage_Current = mLngPages - 1
          'puts random inputs and outputs, but following visual distribution
          lLngTmp1 = Int(Rnd() * mIntORows)
          lLngTmp2 = (mIntICols * mIntIRows - 1) / mIntORows
          'generates inputs
          For j = 1 To 30
            Do
              lLngTmp3 = lLngTmp2 * lLngTmp1 + Int(Rnd() * lLngTmp2)
            Loop Until lLngTmp3 <= SafeUbound(mPages(mLngPage_Current).SngInputs)
            mPages(mLngPage_Current).SngInputs(lLngTmp3) = 1
          Next j
          If lSngTmp1 = 0 Then 'if you want to see how it recognizes certain paterns, for example the first row, add: Or lLngTmp1 = 0
            mPages(mLngPage_Current).SngOutputs(lLngTmp1) = 1
            'if no random error
          ElseIf Rnd() * 100 >= lSngTmp1 Then
            mPages(mLngPage_Current).SngOutputs(lLngTmp1) = 1
            'if random error
          Else
            mPages(mLngPage_Current).SngOutputs(Int(Rnd() * mIntORows)) = 1
          End If
          Call PaintIOs(picInputs, mIntICols, mIntIRows, mPages(mLngPage_Current).SngInputs, False)
          Call PaintIOs(picOutputs, mIntOCols, mIntORows, mPages(mLngPage_Current).SngOutputs, True)
          Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, (mLngPage_Current + 1) & "/" & mLngPages, True, False)
        Next i
      End If
    End If

  End Sub

  Private Sub PageDelete()
    Dim i As Long

    If mLngPage_Current < mLngPages - 1 Then
      For i = mLngPage_Current To mLngPages - 2
        With mPages(i)
          .SngInputs = mPages(i + 1).SngInputs
          .SngOutputs = mPages(i + 1).SngOutputs
        End With
      Next i
    End If

    mLngPages -= 1
    If mLngPage_Current + 1 > mLngPages Then mLngPage_Current = mLngPages - 1
    ReDim Preserve mPages(mLngPages)

  End Sub

  Private Function Outputs(ByVal pBolForceThinking As Boolean) As Single()
    Dim i, j As Long
    Dim lSngOutputs() As Single
    Dim lLngCount As Long

    If mBolThinking Or pBolForceThinking Then
      ReDim lSngOutputs(mIntOCols * mIntORows - 1)
      Call PutInputs(mLngPage_Current)
      For i = 0 To mIntORows - 1
        For j = 0 To mIntOCols - 1
          lSngOutputs(i * mIntOCols + j) = NetOutputGet(lLngCount, 0)
          lLngCount += 1
        Next j
      Next i
      Outputs = lSngOutputs
    Else
      Outputs = mPages(mLngPage_Current).SngOutputs
    End If

  End Function

  Private Sub PagesLearn()
    Dim lDatTmp1 As Date
    Dim lDatTmp3 As Date
    Dim lDatTmp2 As Date
    Dim lDatNow As Date
    Dim lLngTmp As Long
    Dim i, j As Long
    Dim i_prev, j_prev As Long
    Dim lSngPercentOk As Single
    Dim lSngPercentMax As Single
    Dim lSngPercentOk_Prev As Single
    Dim lLngCycleWhereMaxHappened As Long
    Dim lSngPercentageTarget As Single
    Dim lIntTmp As Integer
    Dim lBolUpdatedTimeToFinish As Boolean
    Dim lStrTimeToFinish As String
    Dim lBolEstimateSuccess As Boolean
    Dim lIntAchievedTimes As Integer

    'seconds
    Const TIME_TO_UPDATE_ESTIMATED_FINISHING_TIME = 5
    Const TIME_TO_UPDATE_ESTIMATED_SUCCESS = 3
    Const TIMES_TO_ACHIEVE_SUCCESS_TO_FINISH = 1

    mBolWorking = False 'stops previous trainings

    'creating memory structures
    Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, "Creating memory structures, please be patient...", True, True)
    lDatTmp1 = Now()
    Call LearnPage()
    lLngTmp = DateDiff("s", lDatTmp1, Now())
    lLngTmp = Val(InputBox("Trained 1 page for testing and to generate memory structures. It took " & lLngTmp & " seconds. How many epochs (all pages) do you want to do now?", "Epochs", 0))
    If lLngTmp <> 0 Then
      lSngPercentageTarget = Val(InputBox("Target percentage of success?", "Target percentage", 95))
      If lSngPercentageTarget <> 0 Then
        lSngPercentageTarget /= 100
        mBolWorking = True
        Call NetSetDestroy 'destroys previous sets
        If NetSetPrepare(mLngPages) Then
          For i = 0 To mLngPages - 1
            Call PutInputs(i)
            Call PutOutputs(i)
            If NetSetRecord() <> i + 1 Then
              MsgBox("Error in NetSetRecord", vbCritical)
            End If
            If i Mod 20 = 0 Then
              Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, "Creating set of records " & i + 1 & "/" & mLngPages & "...", True, True)
              If Not mBolWorking Then Exit For
            End If
          Next i
          i_prev = -1
          If mBolWorking Then
            'prepares learning, but keeping the knowledge of previous lessons (pBolNetInitialize = false)
            Call NetSetStart(0, False)
            lDatTmp1 = Now()
            lDatTmp2 = DateAdd(DateInterval.Second, -TIME_TO_UPDATE_PROGRESS - 1, lDatTmp1) 'forces an immediate update on progress
            lStrTimeToFinish = "estimating finish time"
            For i = 0 To lLngTmp - 1
              If mBolWorking Then
                lIntTmp = NetSetLearnStart(PERCENTAGE_OF_ACTIVATION_TO_CONSIDER_ACTIVE, 0, 0)
                If lIntTmp <= NetLearn_NAN Then
                  'estimation of success must be performed for all pages
                  If i = lLngTmp - 1 Then
                    lBolEstimateSuccess = True
                  ElseIf DateDiff(DateInterval.Second, lDatTmp3, lDatNow) > TIME_TO_UPDATE_ESTIMATED_SUCCESS Then
                    lDatTmp3 = lDatNow
                    lBolEstimateSuccess = True
                  End If
                  For j = 0 To mLngPages - 1
                    lDatNow = Now()
                    If DateDiff(DateInterval.Second, lDatTmp1, lDatNow) > TIME_TO_UPDATE_ESTIMATED_FINISHING_TIME And j > j_prev Then
                      lStrTimeToFinish = "to finish: " &
                                         DateAndTime.DateAdd(DateInterval.Second, DateDiff(DateInterval.Second, lDatTmp1, lDatNow) * ((mLngPages - 1 - j) + (lLngTmp - 1 - i) * mLngPages) /
                                         ((i - i_prev) * mLngPages + Math.Abs(j - j_prev)), lDatNow)
                      lDatTmp1 = lDatNow
                      lBolUpdatedTimeToFinish = True
                    End If
                    ' checks estimated rate of success periodically and in the last step of every epoch
                    lSngPercentOk = NetSetLearnContinue(j, lBolEstimateSuccess, False)
                    lDatNow = Now()
                    If DateDiff(DateInterval.Second, lDatTmp2, lDatNow) > TIME_TO_UPDATE_PROGRESS Then
                      lDatTmp2 = lDatNow
                      Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, "Learning (" & lStrTimeToFinish & ") epoch " & i + 1 & "/" & lLngTmp & " (page " & j + 1 & "/" & mLngPages & "): " &
                                       FormatNumber(lSngPercentOk * 100, 2, True, False, True) & "% success, max.: " & FormatNumber(lSngPercentMax * 100, 2, True, False, True) &
                                       "% at epoch " & lLngCycleWhereMaxHappened, True, True)
                    End If
                    If lBolUpdatedTimeToFinish Then
                      i_prev = i
                      j_prev = j
                      lBolUpdatedTimeToFinish = False
                    End If
                    If Not mBolWorking Then
                      Exit For
                    End If
                  Next j
                  lBolEstimateSuccess = False
                  lSngPercentOk = NetSetLearnEnd()
                Else
                  Call LearnResultCheck(lIntTmp)
                  Exit For
                End If
                If lSngPercentOk > lSngPercentMax Then
                  lSngPercentMax = lSngPercentOk
                  Call NetSnapshotTake
                  lLngCycleWhereMaxHappened = i
                End If
                If lSngPercentOk >= lSngPercentageTarget Then
                  lIntAchievedTimes += 1
                  If lIntAchievedTimes >= TIMES_TO_ACHIEVE_SUCCESS_TO_FINISH Then
                    Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, "Learned (" & FormatNumber(lSngPercentMax * 100, 2, True, False, True) & "% success) at epoch " & i + 1 & "/" & lLngTmp, False, False)
                    Exit For
                  End If
                End If
                lSngPercentOk_Prev = lSngPercentOk
              Else
                Exit For
              End If
            Next i
            If lSngPercentOk < lSngPercentageTarget Then
              Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, "Not learned (" & FormatNumber(lSngPercentMax * 100, 2, True, False, True) & "% success) at epoch " & i + 1 & "/" & lLngTmp, False, False)
            Else
              Call CaptionUpdate(Me, mBolThinking, mBolBlackAndWhite, "Learned (" & FormatNumber(lSngPercentMax * 100, 2, True, False, True) & "% success) at epoch " & i + 1 & "/" & lLngTmp, False, False)
            End If
            If lSngPercentMax <> 0 Then
              If Not NetSnapshotGet() Then
                MsgBox("Could not get snapshot from memory.", vbCritical)
              End If
            End If
          End If
        Else
          MsgBox("Could not prepare set of records into memory. Possibly, out of memory.", vbCritical)
        End If
        mBolWorking = False
      End If
    End If

  End Sub

  Private Sub LearnPage()
    Call PutInputs(mLngPage_Current)
    Call PutOutputs(mLngPage_Current)
    LearnResultCheck(NetLearn(0))
  End Sub

  Private Sub LearnResultCheck(ByVal pInt As Integer)
    Select Case pInt
      Case NetLearn_NAN : MsgBox("Not A Number was obtained.", vbExclamation)
      Case NetLearn_ThreadsError : MsgBox("There was an error managing multithreading. Set multithreading to 1 to avoid this error.", vbCritical)
      Case NetLearn_SharedMemoryError : MsgBox("There was an error managing shared memory. This error should not occur in Windows. Set multithreading to 1 to avoid this error.", vbCritical)
      Case NetLearn_SetHasNoRecords : MsgBox("Set has no records to learn.", vbInformation)
    End Select
  End Sub

  Private Sub PutInputs(ByVal pLngPage As Long)
    Dim i As Long
    Dim j As Long
    Dim lLngCount As Long

    For i = 0 To mIntIRows - 1
      For j = 0 To mIntICols - 1
        Call NetInputSet(lLngCount, mPages(pLngPage).SngInputs(i * mIntICols + j))
        lLngCount += 1
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
        lLngCount += 1
      Next j
    Next i

  End Sub

  Private Sub PaintPage()
    Call PaintIOs(picInputs, mIntICols, mIntIRows, mPages(mLngPage_Current).SngInputs, False)
    Call PaintIOs(picOutputs, mIntOCols, mIntORows, Outputs(False), True)
  End Sub

  Private Sub MemorizeIOAndCreateNewPage()
    mLngPages += 1
    Call IOReserveMemoryForThisPage()
    mLngPage_Current = mLngPages - 1
  End Sub

  Private Sub IOClear()
    Dim i As Long

    For i = 0 To mIntICols * mIntIRows - 1
      mPages(mLngPage_Current).SngInputs(i) = 0
    Next i

    For i = 0 To mIntOCols * mIntORows - 1
      mPages(mLngPage_Current).SngOutputs(i) = 0
    Next i

    Call PaintIOs(picInputs, mIntICols, mIntIRows, mPages(mLngPage_Current).SngInputs, False)
    Call PaintIOs(picOutputs, mIntOCols, mIntORows, Outputs(False), True)

  End Sub

  Private Sub IOReserveMemoryForThisPage()

    ReDim Preserve mPages(mLngPages - 1)

    With mPages(mLngPages - 1)
      ReDim .SngInputs(mIntICols * mIntIRows - 1)
      ReDim .SngOutputs(mIntOCols * mIntORows - 1)
    End With

  End Sub

  Private Sub UpdateForm(ByVal pBolCaptionRefresh As Boolean)
    Call ResizeIOCanvas()
    'Call PaintGrid(picInputs, mIntICols, mIntIRows)
    Call PaintGrid(picOutputs, mIntOCols, mIntORows)
    Call PageRefresh(pBolCaptionRefresh)
  End Sub

  Private Sub ResizeIOCanvas()
    Dim lSngTmp As Single

    With picInputs
      .Left = 0
      .Top = 0
      .Width = ClientRectangle.Width * mSngPctIO
      lSngTmp = .Width
      .Height = ClientRectangle.Height
    End With

    With picOutputs
      .Top = 0
      .Left = lSngTmp
      .Width = ClientRectangle.Width - lSngTmp
      .Height = ClientRectangle.Height
    End With

  End Sub

  Private Sub PaintGrid(pPic As PictureBox,
                        ByVal pIntCols As Integer,
                        ByVal pIntRows As Integer)
    Dim i As Integer
    Dim lSngTmp1 As Single
    Dim lSngTmp2 As Single

    Dim lG As Graphics = pPic.CreateGraphics
    Dim lPen As New Drawing.Pen(Color.FromArgb(mLngColorGrey, mLngColorGrey, mLngColorGrey))

    lG.Clear(Color.FromArgb(255, 255, 255))

    lSngTmp1 = pPic.Width / pIntCols

    For i = 1 To pIntCols - 1
      lSngTmp2 = lSngTmp1 * i
      lG.DrawLine(lPen, lSngTmp2, 0, lSngTmp2, pPic.Height)
    Next i

    lSngTmp1 = pPic.Height / pIntRows

    For i = 1 To pIntRows - 1
      lSngTmp2 = lSngTmp1 * i
      lG.DrawLine(lPen, 0, lSngTmp2, pPic.Width, lSngTmp2)
    Next i

  End Sub

  Private Sub PaintIOs(pPic As PictureBox,
                       ByVal pIntCols As Integer,
                       ByVal pIntRows As Integer,
                       pSngIO() As Single,
                       ByVal pBolOutput As Boolean)
    Dim i, j As Integer
    Dim lSngTmp1 As Single
    Dim lSngTmp2 As Single

    With pPic
      lSngTmp1 = .Width / pIntCols
      lSngTmp2 = .Height / pIntRows
    End With

    For i = 0 To pIntRows - 1
      For j = 0 To pIntCols - 1
        Call PaintIO(pPic, pIntCols, pIntRows, j, i, pSngIO, pBolOutput)
      Next j
    Next i

  End Sub

  Private Sub PaintIO(pPic As PictureBox,
                      ByVal pIntCols As Integer,
                      ByVal pIntRows As Integer,
                      ByVal pIntCol As Integer,
                      ByVal pIntRow As Integer,
                      pSngIO() As Single,
                      ByVal pBolOutput As Boolean)
    Dim lSngTmp1 As Single
    Dim lSngTmp2 As Single
    Dim lLngTmp As Long
    Dim lLngIdx As Long
    Dim lIntRed As Integer
    Dim lIntGreen As Integer
    Dim lIntBlue As Integer
    Dim lIntColorAvg As Integer
    Dim lColor As New Color

    With pPic
      lSngTmp1 = .Width / pIntCols
      lSngTmp2 = .Height / pIntRows
      lLngIdx = pIntRow * pIntCols + pIntCol

      If CStr(pSngIO(lLngIdx)) <> "1,#QNAN" Then
        lLngTmp = MAX_COLOR * (1 - Val(Replace$(pSngIO(lLngIdx), ",", ".")))
      Else
        lLngTmp = MAX_COLOR
      End If

      If pBolOutput Then
        'B&W based on activation
        If mBolBlackAndWhite Then
          If lLngTmp > MAX_COLOR * PERCENTAGE_OF_ACTIVATION_TO_CONSIDER_ACTIVE Then lLngTmp = MAX_COLOR Else lLngTmp = 0
        End If
        lColor = ColorTranslator.FromOle(lLngTmp)
        'For grey tones
      ElseIf mBolBlackAndWhite Then
        lIntRed = ColorTranslator.FromOle(lLngTmp).R
        lIntGreen = ColorTranslator.FromOle(lLngTmp).G
        lIntBlue = ColorTranslator.FromOle(lLngTmp).B
        lIntColorAvg = (lIntRed + lIntGreen + lIntBlue) / 3
        lColor = Color.FromArgb(lIntColorAvg, lIntColorAvg, lIntColorAvg)
      Else
        lColor = ColorTranslator.FromOle(lLngTmp)
      End If

      Dim lG As Graphics = .CreateGraphics
      Dim lRect As New Rectangle(lSngTmp1 * pIntCol, lSngTmp2 * pIntRow, lSngTmp1, lSngTmp2)
      Dim lBrush As New SolidBrush(lColor)
      'FillStyle = solid
      lG.FillRectangle(lBrush, lRect)
      If pBolOutput Then
        Dim lPen As New Drawing.Pen(Color.FromArgb(mLngColorGrey, mLngColorGrey, mLngColorGrey))
        lG.DrawRectangle(lPen, lSngTmp1 * pIntCol, lSngTmp2 * pIntRow, lSngTmp1, lSngTmp2)
      End If
    End With

  End Sub

  Private Sub InputsMouseDown(Button As MouseButtons, X As Single, Y As Single)
    Dim lIntCol As Integer
    Dim lIntRow As Integer
    Dim lBolTmp As Boolean

    If Button = MouseButtons.Left Then
      With picInputs
        lIntCol = Int(X / (.Width / mIntICols))
        lIntRow = Int(Y / (.Height / mIntIRows))
      End With
      lBolTmp = lIntCol >= 0 And lIntRow >= 0 And lIntCol < mIntICols And lIntRow < mIntIRows
    ElseIf Button = MouseButtons.Right Then
      Call ShowHelp()
    End If

    If lBolTmp Then
      mPages(mLngPage_Current).SngInputs(lIntRow * mIntICols + lIntCol) = 1 - mLngColor / MAX_COLOR
      Call PaintIO(picInputs, mIntICols, mIntIRows, lIntCol, lIntRow, mPages(mLngPage_Current).SngInputs, False)
      If mBolThinking Then Call PaintIOs(picOutputs, mIntOCols, mIntORows, Outputs(False), True)
    End If

  End Sub

  Private Sub OutputsMouseDown(Button As MouseButtons, X As Single, Y As Single)
    Dim lIntCol As Integer
    Dim lIntRow As Integer
    Dim lBolTmp As Boolean

    If Not mBolThinking Then
      If Button = MouseButtons.Left Then
        With picOutputs
          lIntCol = Int(X / (.Width / mIntOCols))
          lIntRow = Int(Y / (.Height / mIntORows))
        End With
        lBolTmp = lIntCol >= 0 And lIntRow >= 0 And lIntCol < mIntOCols And lIntRow < mIntORows
      ElseIf Button = MouseButtons.Right Then
        Call ShowHelp()
      End If

      If lBolTmp Then
        mPages(mLngPage_Current).SngOutputs(lIntRow * mIntOCols + lIntCol) = 1 - mLngColor / MAX_COLOR
        Call PaintIO(picOutputs, mIntOCols, mIntORows, lIntCol, lIntRow, mPages(mLngPage_Current).SngOutputs, True)
      End If
    End If

  End Sub

  Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
    If Not mBolActivated Then
      Cursor.Current = Cursors.WaitCursor
      mBolActivated = InitializeAutomatically()
      If mBolActivated Then Call ShowHelp() Else Me.Close()
    End If
  End Sub

  Private Sub Form1_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
    mSngPctIO = e.X / Me.Width
    Call ResizeIOCanvas()
    'Call PaintGrid(picInputs, mIntICols, mIntIRows)
    Call PaintGrid(picOutputs, mIntOCols, mIntORows)
    Call PaintIOs(picInputs, mIntICols, mIntIRows, mPages(mLngPage_Current).SngInputs, False)
    Call PaintIOs(picOutputs, mIntOCols, mIntORows, Outputs(False), True)
  End Sub

  Private Sub picInputs_MouseDown(sender As Object, e As MouseEventArgs) Handles picInputs.MouseDown
    Call InputsMouseDown(e.Button, e.X, e.Y)
  End Sub

  Private Sub picInputs_MouseMove(sender As Object, e As MouseEventArgs) Handles picInputs.MouseMove
    Call InputsMouseDown(e.Button, e.X, e.Y)
  End Sub

  Private Sub picOutputs_MouseDown(sender As Object, e As MouseEventArgs) Handles picOutputs.MouseDown
    Call OutputsMouseDown(e.Button, e.X, e.Y)
  End Sub

  Private Sub picOutputs_MouseMove(sender As Object, e As MouseEventArgs) Handles picOutputs.MouseMove
    Call OutputsMouseDown(e.Button, e.X, e.Y)
  End Sub

  Private Sub Form1_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
    If mBolActivated Then Call UpdateForm(False)
  End Sub

  Private Sub Form1_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
    If mBolActivated Then Call UpdateForm(False)
  End Sub

  Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    mBolWorking = False
    Application.DoEvents()
    Call NetDestroy()
  End Sub
End Class
