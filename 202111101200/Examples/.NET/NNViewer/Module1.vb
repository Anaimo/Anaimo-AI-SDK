Module Module1

  Public Const FORM_CAPTION = "NNViewer v1"
  Public Const FORM_CAPTION_SEPARATOR = " - "

  Private mLngFinalNumberOfTotalNeurons As Long

  Public Structure tIOPage
    Dim SngInputs() As Single
    Dim SngOutputs() As Single
  End Structure

  '---------------------------------------------------------
  '(c)2004-2022, Anaimo (R) AI Technology declarations
  '---------------------------------------------------------
  Public Const NetCreate_Success = 0
  Public Const NetCreate_LicenseExpiresInLessThan30Days = 1
  Public Const NetCreate_NotLicensed = 2
  Public Const NetCreate_OutOfMemory = 3
  Public Const NetCreate_UnknownError = 4

  Public Const NetLearn_Success = 0
  Public Const NetLearn_NAN = 1
  Public Const NetLearn_ThreadsError = 2
  Public Const NetLearn_SharedMemoryError = 3
  Public Const NetLearn_SetHasNoRecords = 4

  Public Const MODE_STANDARD_BACKPROPAGATION = 0
  Public Const MODE_STANDARD_BACKPROP_OPTIMIZED = 1
  Public Const MODE_DYNAMIC_PROPAGATION = 2
  Public Const MODE_FINISH = MODE_DYNAMIC_PROPAGATION 'to define the end of modes

  Declare Function HardwareId Lib "AnaimoAI.dll" (ByVal pIntRow As Long, ByVal pIntCol As Long) As Long
  Declare Sub NetActivationFunctionSet Lib "AnaimoAI.dll" (ByVal pLng As Long)
  Declare Function NetCreate Lib "AnaimoAI.dll" (ByVal pLngMaxNeurons As Long,
                                              ByVal pLngMaxInputs As Long,
                                              ByVal pLngMaxOutputs As Long) As Long
  Declare Sub NetDestroy Lib "AnaimoAI.dll" ()
  Declare Function NetErrorGet Lib "AnaimoAI.dll" (ByVal pLngCyclesControl As Long) As Single
  Declare Sub NetInitialize Lib "AnaimoAI.dll" (ByVal pSngDefaultVal As Single)
  Declare Function NeuBiasGet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long) As Single
  Declare Function NeuDeltaGet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long) As Single
  Declare Function NeuValueGet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long) As Single
  Declare Function NeuValueUpdatedGet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long) As Long
  Declare Function NeuInputWeightGet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long, ByVal pLngInput As Long) As Single
  Declare Function NeuInputWeightUpdatedGet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long, ByVal pLngInput As Long) As Long
  Declare Function NeuInputsNumberGet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long) As Long
  Declare Sub NeuBiasSet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long, ByVal pSngVal As Single)
  Declare Sub NeuInputWeightSet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long, ByVal pLngInput As Long, ByVal pSngVal As Single)
  Declare Sub NetOutputAdd Lib "AnaimoAI.dll" (ByVal pLngSrc As Long)
  Declare Sub NetNeuronAdd Lib "AnaimoAI.dll" ()
  Declare Function NetNeuronsAddedNumberGet Lib "AnaimoAI.dll" () As Long
  Declare Function NetInputsMaxNumberGet Lib "AnaimoAI.dll" () As Long
  Declare Function NetNeuronsMaxNumberGet Lib "AnaimoAI.dll" () As Long
  Declare Function NetConnect Lib "AnaimoAI.dll" (ByVal pLngSrc As Long, ByVal pLngDst As Long) As Boolean
  Declare Function NetConnectConsecutive Lib "AnaimoAI.dll" (ByVal pLngSrc1 As Long, ByVal pLngSrc2 As Long, ByVal pLngDst As Long) As Boolean
  Declare Function NetDropOutGet Lib "AnaimoAI.dll" () As Single
  Declare Sub NetDropOutSet Lib "AnaimoAI.dll" (ByVal pSng As Single)
  Declare Sub NetInputSet Lib "AnaimoAI.dll" (ByVal pLngInput As Long, ByVal pSngVal As Single)
  Declare Function NetMomentumGet Lib "AnaimoAI.dll" () As Single
  Declare Sub NetMomentumSet Lib "AnaimoAI.dll" (ByVal pSng As Single)
  Declare Function NetOutputGet Lib "AnaimoAI.dll" (ByVal pLngOutput As Long, ByVal pLngCyclesControl As Long) As Single
  Declare Function NetOutputsAddedNumberGet Lib "AnaimoAI.dll" () As Long
  Declare Function NetOutputsMaxNumberGet Lib "AnaimoAI.dll" () As Long
  Declare Sub NetOutputSet Lib "AnaimoAI.dll" (ByVal pLngOutput As Long, ByVal pSngVal As Single)
  Declare Function NetLearningRateGet Lib "AnaimoAI.dll" () As Single
  Declare Sub NetLearningRateSet Lib "AnaimoAI.dll" (ByVal pSng As Single)
  Declare Function NetLearn Lib "AnaimoAI.dll" (ByVal pLngCyclesControl As Long) As Integer
  Declare Function NetModeGet Lib "AnaimoAI.dll" () As Long
  Declare Sub NetModeSet Lib "AnaimoAI.dll" (ByVal pInt As Long)
  Declare Sub NetSetDestroy Lib "AnaimoAI.dll" ()
  'call NetSetStart after having NetSetRecord all records
  Declare Function NetSetStart Lib "AnaimoAI.dll" (ByVal pSngDefaultVal As Single, ByVal pBolNetInitialize As Boolean) As Boolean
  Declare Function NetSetPrepare Lib "AnaimoAI.dll" (ByVal pLngTotalNumberOfRecords As Long) As Boolean
  Declare Function NetSetRecord Lib "AnaimoAI.dll" () As Long
  Declare Function NetSetLearnStart Lib "AnaimoAI.dll" (ByVal pSngThresholdForActive As Single,
                                                     ByVal pSngDeviationPercentageTarget As Single,
                                                     ByVal pLngCyclesControl As Long) As Integer
  Declare Function NetSetLearnContinue Lib "AnaimoAI.dll" (ByVal pLngRecordNumber As Long, ByVal pBolEstimateSuccess As Boolean, ByRef pBolThereIsNan As Boolean) As Single
  Declare Function NetSetLearnEnd Lib "AnaimoAI.dll" () As Single
  Declare Function NetSnapshotTake Lib "AnaimoAI.dll" () As Long
  Declare Function NetSnapshotGet Lib "AnaimoAI.dll" () As Boolean
  Declare Function NetThreadsMaxNumberGet Lib "AnaimoAI.dll" () As Long
  Declare Sub NetThreadsMaxNumberSet Lib "AnaimoAI.dll" (ByVal pInt As Long)

  Public Function SafeUbound(pArray As Object,
                             Optional ByVal pIntDim As Integer = 0) As Long

    Try
      If pIntDim > 1 Then
        SafeUbound = UBound(pArray, pIntDim)
      Else
        SafeUbound = UBound(pArray)
      End If
    Catch ex As Exception
      Return (-1)
    End Try

  End Function

  'adds an item a tLngPieces list
  Public Sub AddToLngList(pPieces() As Long, ByVal pLngItem As Long)
    ReDim Preserve pPieces(SafeUbound(pPieces) + 1)
    pPieces(SafeUbound(pPieces)) = pLngItem
  End Sub

  Private Function MyNetCreate(ByVal pLngMaxNeurons As Long,
                               ByVal pLngMaxInputs As Long,
                               ByVal pLngMaxOutputs As Long) As Long
    Dim lLngRes As Long

    lLngRes = NetCreate(pLngMaxNeurons, pLngMaxInputs, pLngMaxOutputs)

    If lLngRes = NetCreate_LicenseExpiresInLessThan30Days Or lLngRes = NetCreate_NotLicensed Then
      If lLngRes = NetCreate_LicenseExpiresInLessThan30Days Then
        MsgBox("Neural Network license will expire in less than 30 days. Please request your new licensed version by providing to Anaimo this hardware id (you can paste it as has been copied to the clipboard): " & MyHardwareIdAndIntoClipBoard(), vbExclamation)
      ElseIf lLngRes = NetCreate_NotLicensed Then
        MsgBox("Neural Network license is not present or expired. Please request your licensed version by providing to Anaimo this hardware id (you can paste it as has been copied to the clipboard): " & MyHardwareIdAndIntoClipBoard(), vbExclamation)
      Else
        MsgBox("Could not initialize the neural network (4). Error: " & lLngRes, vbCritical)
      End If
    End If

    Return lLngRes

  End Function

  Public Function CreateNeuralNetwork(pForm As Form,
                                      ByVal pLngMaxNumberOfNeurons As Long,
                                      ByVal pLngInputs As Long,
                                      ByVal pLngOutputs As Long,
                                      ByVal pBytConnectHow As Byte,
                                      ByVal pIntMaxNumberOfLayers As Integer,
                                      ByRef pIntNumberOfLayers As Integer,
                                      ByRef pStrLayersBounds As String) As Boolean
    Const PROC_NAME = "CreateNeuralNetwork"
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim X As Long
    Dim l As Long
    Dim lLngCount As Long
    Dim lLngLastLayerNumberOfNeurons As Long
    Dim lIntLayerNum As Integer
    Dim lLngRes As Long
    'to connect to later layers
    Dim lNeuronNumber() As Long = Nothing
    Dim lLngPreviousLayers As Long
    Dim lLayerNeuronsStartIndex() As Long = Nothing
    Dim lLayerNeuronsEndIndex() As Long = Nothing
    Dim lLngStepMax As Long
    Dim lLngStep As Long

    'random connections
    Dim lLngSrc As Long
    Dim lLngDst As Long

    Const RANDOM_CONNECTIONS = 1000000

    Const UPDATE_STATUS_EVERY = 100

    Try
      'initialize
      pIntNumberOfLayers = 0
      pStrLayersBounds = vbNullString
      mLngFinalNumberOfTotalNeurons = 0
      Call NetDestroy

      If pLngOutputs > pLngInputs Then
        MsgBox("Neural network cannot be created as inputs (" & pLngInputs & ") must be greater or equal than outputs (" & pLngOutputs & ").", vbCritical)
      ElseIf pLngInputs = pLngOutputs Then '(WARNING: not tested intensively)
        pIntNumberOfLayers = pIntMaxNumberOfLayers
        'calculates number of neurons
        mLngFinalNumberOfTotalNeurons = pLngInputs * pIntMaxNumberOfLayers
        If mLngFinalNumberOfTotalNeurons >= pLngMaxNumberOfNeurons Then
          MsgBox("Maximum number of neurons reached.", vbCritical)
        Else
          lLngRes = MyNetCreate(mLngFinalNumberOfTotalNeurons, pLngInputs, pLngOutputs)
          If lLngRes <= NetCreate_LicenseExpiresInLessThan30Days Then
            pStrLayersBounds = "Layer 1: 0->" & (pLngInputs - 1) & vbCrLf
            For k = 1 To pIntMaxNumberOfLayers
              For i = 1 To pLngInputs
                Call NetNeuronAdd
                lLngCount += 1
              Next i
            Next k

            'Sets the outputs (on neurons already created)
            For i = 0 To pLngOutputs - 1
              Call NetOutputAdd(mLngFinalNumberOfTotalNeurons - pLngOutputs + i)
            Next i

            'calculates number of steps doing the same loops
            For k = 2 To pIntMaxNumberOfLayers
              For j = 0 To pLngInputs - 1
                lLngStepMax += 1
              Next
            Next

            lLngCount = 0
            'neuron connections
            For k = 2 To pIntMaxNumberOfLayers
              pStrLayersBounds = pStrLayersBounds & "Layer " & k & ": " & pLngInputs * (k - 1) & "->" & pLngInputs * k - 1 & vbCrLf
              For j = 0 To pLngInputs - 1
                'Debug.Print "Connecting " & i & "->" & pLngInputs * (k - 1) + j
                If Not NetConnectConsecutive(pLngInputs * (k - 2), pLngInputs * (k - 1) - 1, pLngInputs * (k - 1) + j) Then
                  MsgBox("Could not initialize the neural network (1).", vbCritical)
                  lIntLayerNum = -1 'to exit
                  Exit For
                End If
                If lLngStep Mod UPDATE_STATUS_EVERY = 0 Then
                  Call CaptionUpdate(pForm, False, False, "Creating neural network (" & FormatNumber(lLngStep / lLngStepMax * 100, 2) & "%)...", True, False)
                End If
                lLngStep += 1
              Next j
              If lIntLayerNum = -1 Then Exit For
            Next k
            pStrLayersBounds = pStrLayersBounds & "Outputs: " & mLngFinalNumberOfTotalNeurons - pLngOutputs & "->" & mLngFinalNumberOfTotalNeurons - 1
            'success?
            Return lIntLayerNum <> -1
          End If
        End If
      Else
        X = (pLngInputs - pLngOutputs) / (pIntMaxNumberOfLayers - 1) 'substracts 1 for the outputs
        If X < 1 Then
          MsgBox("Neural network cannot be created as reduction in between layers (" & X & ") must be greater than 1.", vbCritical)
        Else
          'inputs are considered like neurons
          mLngFinalNumberOfTotalNeurons = pLngInputs
          'calculates number of neurons
          For k = pLngInputs To pLngOutputs + X Step -X
            pIntNumberOfLayers += 1
            mLngFinalNumberOfTotalNeurons += k
            lLngLastLayerNumberOfNeurons = k
          Next k
          'pIntNumberOfLayers is now +1 because it includes the outputs
          If mLngFinalNumberOfTotalNeurons >= pLngMaxNumberOfNeurons Then
            MsgBox("Maximum number of neurons reached.", vbCritical)
          ElseIf lLngLastLayerNumberOfNeurons <= pLngOutputs Then
            MsgBox("Cannot add the neurons for the outputs as the number of neurons in the last layer is less or equal than the number of needed outputs.", vbCritical)
          Else
            'outputs are also considered neurons
            mLngFinalNumberOfTotalNeurons += pLngOutputs
            If mLngFinalNumberOfTotalNeurons >= pLngMaxNumberOfNeurons Then
              MsgBox("Total number of neurons (" & mLngFinalNumberOfTotalNeurons & ") exceeds the maximum number (" & pLngMaxNumberOfNeurons & ") of possible neurons.", vbCritical)
            Else
              lLngRes = MyNetCreate(mLngFinalNumberOfTotalNeurons, pLngInputs, pLngOutputs)
              If lLngRes <= NetCreate_LicenseExpiresInLessThan30Days Then

                lLngCount = 0
                pStrLayersBounds = "Inputs: 0->" & (pLngInputs - 1) & vbCrLf

                'adds inputs (by default, the first number of inputs neurons are inputs)
                For k = 1 To pLngInputs
                  Call NetNeuronAdd
                  lLngCount += 1
                Next

                pStrLayersBounds = pStrLayersBounds & "Layer 1: " & pLngInputs & "->" & (pLngInputs * 2 - 1) & vbCrLf

                'same FOR than above (*)
                For k = pLngInputs To pLngOutputs + X Step -X
                  For i = 1 To k
                    Call NetNeuronAdd
                    lLngCount += 1
                  Next i
                Next k

                'adds one more layer for the outputs
                pIntNumberOfLayers += 1

                'adds outputs' neurons
                For i = 1 To pLngOutputs
                  Call NetNeuronAdd
                  lLngCount += 1
                Next i

                'tells the NN which neurons are the outputs (on neurons already created)
                For i = 0 To pLngOutputs - 1
                  Call NetOutputAdd(lLngCount - pLngOutputs + i)
                Next i

                'random connections (WARNING: not tested intensively)
                If pBytConnectHow = 2 Then
                  Randomize()
                  pStrLayersBounds = pStrLayersBounds & RANDOM_CONNECTIONS & " connections" & vbCrLf
                  For i = 0 To RANDOM_CONNECTIONS
                    Do
                      lLngSrc = (mLngFinalNumberOfTotalNeurons - pLngOutputs) * Rnd()
                      lLngDst = mLngFinalNumberOfTotalNeurons * Rnd()
                    Loop Until lLngSrc <> lLngDst
                    If Not NetConnect(lLngSrc, lLngDst) Then
                      MsgBox("Could not initialize the neural network (1).", vbCritical)
                      lIntLayerNum = -1 'to exit
                      Exit For
                    End If
                  Next i
                  lLngCount = 0
                  'shows layers
                  For k = pLngInputs To pLngOutputs + X Step -X
                    If lLngCount + 2 * k - X - 1 < mLngFinalNumberOfTotalNeurons Then
                      pStrLayersBounds = pStrLayersBounds & "Layer " & lIntLayerNum + 1 & ": " & (lLngCount + k) & "->" & (lLngCount + 2 * k - X - 1) & vbCrLf
                    End If
                  Next k
                  'CONNECTIONS of every neuron to all later layers (WARNING: not tested intensively)
                ElseIf pBytConnectHow = 1 Then

                  'calculates number of steps doing the same loops
                  For k = pLngInputs To pLngOutputs + X + 1 Step -X '+1 is to avoid doing the output layer here
                    'For i = lLngCount To lLngCount + k - 1
                    For j = lLngCount + k To lLngCount + 2 * k - X - 1
                      lLngStepMax += 1
                    Next
                    'Next
                    For l = 0 To lLngPreviousLayers - 1
                      For i = lLayerNeuronsStartIndex(l) To lLayerNeuronsEndIndex(l)
                        For j = lLngCount + k To lLngCount + 2 * k - X - 1
                          lLngStepMax += 1
                        Next
                      Next
                    Next
                  Next

                  lLngCount = 0
                  lIntLayerNum = 1
                  'neuron connections
                  For k = pLngInputs To pLngOutputs + X + 1 Step -X  '+1 is to avoid doing the output layer here
                    If lLngCount + 2 * k - X - 1 < mLngFinalNumberOfTotalNeurons Then
                      pStrLayersBounds = pStrLayersBounds & "Layer " & lIntLayerNum + 1 & ": " & (lLngCount + k) & "->" & (lLngCount + 2 * k - X - 1) & vbCrLf
                      lIntLayerNum += 1
                      AddToLngList(lLayerNeuronsStartIndex, SafeUbound(lNeuronNumber) + 1)

                      For i = lLngCount To lLngCount + k - 1
                        For j = lLngCount + k To lLngCount + 2 * k - X - 1
                          If Not NetConnect(i, j) Then
                            MsgBox("Could not initialize the neural network (1).", vbCritical)
                            lIntLayerNum = -1 'to exit
                            Exit For
                          Else
                            AddToLngList(lNeuronNumber, i) 'when was here, although incorrect, produced better results
                          End If
                          If lLngStep Mod UPDATE_STATUS_EVERY = 0 Then
                            Call CaptionUpdate(pForm, False, False, "Creating neural network (" & FormatNumber(lLngStep / lLngStepMax * 100, 2) & "%)...", True, False)
                          End If
                          lLngStep += 1
                        Next j
                        If lIntLayerNum = -1 Then
                          Exit For
                        Else
                          AddToLngList(lNeuronNumber, i)
                        End If
                      Next i

                      For i = lLngCount To lLngCount + k - 1
                        If lIntLayerNum = -1 Then
                          Exit For
                        Else
                          AddToLngList(lNeuronNumber, i)
                        End If
                      Next i

                      AddToLngList(lLayerNeuronsEndIndex, SafeUbound(lNeuronNumber))

                      'connects all previous layers to output layer
                      For l = 0 To lLngPreviousLayers - 1
                        For i = lLayerNeuronsStartIndex(l) To lLayerNeuronsEndIndex(l)
                          For j = lLngCount + k To lLngCount + 2 * k - X - 1
                            If Not NetConnect(lNeuronNumber(i), j) Then
                              MsgBox("Could not initialize the neural network (1).", vbCritical)
                              lIntLayerNum = -1 'to exit
                              Exit For
                            End If
                            If lLngStep Mod UPDATE_STATUS_EVERY = 0 Then
                              Call CaptionUpdate(pForm, False, False, "Creating neural network (" & FormatNumber(lLngStep / lLngStepMax * 100, 2) & "%)...", True, False)
                            End If
                            lLngStep += 1
                          Next j
                          If lIntLayerNum = -1 Then Exit For
                        Next i
                      Next l
                      lLngPreviousLayers += 1
                    End If
                    If lIntLayerNum = -1 Then Exit For
                    lLngCount += k
                  Next k

                  'if no error
                  If lIntLayerNum <> -1 Then
                    pStrLayersBounds = pStrLayersBounds & "Outputs: " & mLngFinalNumberOfTotalNeurons - pLngOutputs & "->" & mLngFinalNumberOfTotalNeurons - 1

                    'connects last layer to pLngOutputs
                    For j = mLngFinalNumberOfTotalNeurons - pLngOutputs To mLngFinalNumberOfTotalNeurons - 1
                      If Not NetConnectConsecutive(mLngFinalNumberOfTotalNeurons - pLngOutputs - lLngLastLayerNumberOfNeurons, mLngFinalNumberOfTotalNeurons - pLngOutputs - 1, j) Then
                        MsgBox("Could not initialize the neural network (2).", vbCritical)
                        lIntLayerNum = -1 'to exit
                        Exit For
                      End If
                      If lLngStep Mod UPDATE_STATUS_EVERY = 0 Then
                        Call CaptionUpdate(pForm, False, False, "Creating neural network (" & FormatNumber(lLngStep / lLngStepMax * 100, 2) & "%)...", True, False)
                      End If
                      lLngStep += 1
                    Next j

                    'connects all previous layers to output layer
                    For l = 0 To lLngPreviousLayers - 1
                      For i = lLayerNeuronsStartIndex(l) To lLayerNeuronsEndIndex(l)
                        For j = mLngFinalNumberOfTotalNeurons - pLngOutputs To mLngFinalNumberOfTotalNeurons - 1
                          If Not NetConnect(lNeuronNumber(i), j) Then
                            MsgBox("Could not initialize the neural network (2).", vbCritical)
                            lIntLayerNum = -1 'to exit
                            Exit For
                          End If
                          If lLngStep Mod UPDATE_STATUS_EVERY = 0 Then
                            Call CaptionUpdate(pForm, False, False, "Creating neural network (" & FormatNumber(lLngStep / lLngStepMax * 100, 2) & "%)...", True, False)
                          End If
                          lLngStep += 1
                        Next j
                        If lIntLayerNum = -1 Then Exit For
                      Next i
                    Next l
                  End If
                  ' CONNECTIONS only with next layer
                Else

                  'first connects inputs to layer with the same number of neurons
                  For k = 0 To pLngInputs - 1
                    lLngStepMax += 1
                  Next

                  'calculates number of steps doing the same loops
                  For k = pLngInputs To pLngOutputs + X + 1 Step -X '+1 is to avoid doing the output layer here
                    For j = lLngCount + k To lLngCount + 2 * k - X - 1
                      lLngStepMax += 1
                    Next
                  Next

                  'then connects outputs
                  For j = mLngFinalNumberOfTotalNeurons - pLngOutputs To mLngFinalNumberOfTotalNeurons - 1
                    lLngStepMax += 1
                  Next

                  'after having calculated the number of steps, now it really connects neurons

                  lLngCount = 0
                  lIntLayerNum = 0

                  'first connects inputs to layer with the same number of neurons
                  For k = 0 To pLngInputs - 1
                    If Not NetConnectConsecutive(0, pLngInputs - 1, pLngInputs + k) Then
                      MsgBox("Could not initialize the neural network (1).", vbCritical)
                      lIntLayerNum = -1 'to exit
                      Exit For
                    End If
                    If lLngStep Mod 100 = 0 Then
                      Call CaptionUpdate(pForm, False, False, "Creating neural network (" & FormatNumber(lLngStep / lLngStepMax * 100, 2) & "%)...", True, False)
                    End If
                    lLngStep += 1
                  Next

                  lLngCount = pLngInputs
                  lIntLayerNum += 1
                  'neuron connections
                  For k = pLngInputs To pLngOutputs + X + 1 Step -X '+1 is to avoid doing the output layer here
                    If lLngCount + 2 * k - X - 1 < mLngFinalNumberOfTotalNeurons Then
                      pStrLayersBounds = pStrLayersBounds & "Layer " & lIntLayerNum + 1 & ": " & (lLngCount + k) & "->" & (lLngCount + 2 * k - X - 1) & vbCrLf
                      lIntLayerNum += 1
                      For j = lLngCount + k To lLngCount + 2 * k - X - 1
                        If Not NetConnectConsecutive(lLngCount, lLngCount + k - 1, j) Then
                          MsgBox("Could not initialize the neural network (1).", vbCritical)
                          lIntLayerNum = -1 'to exit
                          Exit For
                        End If
                        If lLngStep Mod 100 = 0 Then
                          Call CaptionUpdate(pForm, False, False, "Creating neural network (" & FormatNumber(lLngStep / lLngStepMax * 100, 2) & "%)...", True, False)
                        End If
                        lLngStep += 1
                      Next j
                    End If
                    If lIntLayerNum = -1 Then Exit For
                    lLngCount += k
                  Next k
                  'if no error
                  If lIntLayerNum <> -1 Then
                    pStrLayersBounds = pStrLayersBounds & "Outputs: " & mLngFinalNumberOfTotalNeurons - pLngOutputs & "->" & mLngFinalNumberOfTotalNeurons - 1
                    'connects last layer to pLngOutputs
                    For j = mLngFinalNumberOfTotalNeurons - pLngOutputs To mLngFinalNumberOfTotalNeurons - 1
                      If Not NetConnectConsecutive(mLngFinalNumberOfTotalNeurons - pLngOutputs - lLngLastLayerNumberOfNeurons, mLngFinalNumberOfTotalNeurons - pLngOutputs - 1, j) Then
                        MsgBox("Could not initialize the neural network (2).", vbCritical)
                        lIntLayerNum = -1 'to exit
                        Exit For
                      End If
                      If lLngStep Mod UPDATE_STATUS_EVERY = 0 Then
                        Call CaptionUpdate(pForm, False, False, "Creating neural network (" & FormatNumber(lLngStep / lLngStepMax * 100, 2) & "%)...", True, False)
                      End If
                      lLngStep += 1
                    Next j
                  End If
                End If
                Return lIntLayerNum <> -1
              End If
            End If
          End If
        End If
      End If
    Catch ex As Exception
      MsgBox("Error in " & PROC_NAME & ": " & ex.Message, vbCritical)
    End Try

    Return False

  End Function

  Public Function MyHardwareId() As String
    Dim i As Long
    Dim j As Long
    Dim lStrTmp As String = ""

    For i = 0 To 3
      For j = 0 To 3
        lStrTmp = lStrTmp & Hex$(HardwareId(i, j)) & ":"
      Next j
    Next i

    Return Strings.Left$(lStrTmp, Strings.Len(lStrTmp) - 1)

  End Function

  Public Function MyHardwareIdAndIntoClipBoard() As String
    Dim lStrHID As String

    lStrHID = MyHardwareId()

    Try
      Clipboard.Clear()
      Clipboard.SetText(lStrHID)
    Catch ex As Exception
      'ignore errors as sometimes Clipboard might not be available
    End Try

    Return lStrHID

  End Function

  Public Sub CaptionUpdate(pForm As Form,
                           ByVal pBolThinking As Boolean,
                           ByVal pBolBlackAndWhite As Boolean,
                           ByVal pStrCaption As String,
                           ByVal pBolRefresh As Boolean,
                           ByVal pBolDoEvents As Boolean)

    If Len(pStrCaption) <> 0 Then
      pForm.Text = FORM_CAPTION & FORM_CAPTION_SEPARATOR & pStrCaption
    Else
      pForm.Text = FORM_CAPTION
    End If

    If pBolBlackAndWhite Then pForm.Text = pForm.Text & FORM_CAPTION_SEPARATOR & "B&W"
    If pBolThinking Then pForm.Text = pForm.Text & FORM_CAPTION_SEPARATOR & "Thinking"

    pForm.Text = pForm.Text & FORM_CAPTION_SEPARATOR & "Type m for help"

    If pBolRefresh Then pForm.Refresh()
    If pBolDoEvents Then Application.DoEvents()

  End Sub

End Module