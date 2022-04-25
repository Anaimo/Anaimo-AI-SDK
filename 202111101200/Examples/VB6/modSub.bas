Attribute VB_Name = "modSub"
Option Explicit

Private mLngFinalNumberOfTotalNeurons As Long
Private mLngInputs As Long

Public Type tIOPage
  SngInputs() As Single
  SngOutputs() As Single
End Type

'---------------------------------------------------------
'(c)2004-2022, Anaimo (R) AI Technology declarations
'---------------------------------------------------------

Public Const NetCreate_Success = 0
Public Const NetCreate_LicenseExpiresInLessThan30Days = 1
Public Const NetCreate_NotLicensed = 2
Public Const NetCreate_OutOfMemory = 3
Public Const NetCreate_UnknownError = 4
  
Declare Function HardwareId Lib "AnaimoAI.dll" (ByVal pIntRow As Long, ByVal pIntCol As Long) As Long
Declare Sub NetActivationFunctionSet Lib "AnaimoAI.dll" (ByVal pLng As Long)
Declare Function NetCreate Lib "AnaimoAI.dll" (ByVal pLngMaxNeurons As Long, _
                                              ByVal pLngMaxInputs As Long, _
                                              ByVal pLngMaxOutputs As Long) As Boolean
Declare Sub NetDestroy Lib "AnaimoAI.dll" ()
Declare Function NetErrorGet Lib "AnaimoAI.dll" (ByVal pLngCyclesControl As Long) As Single
Declare Sub NetInitialize Lib "AnaimoAI.dll" (ByVal pSngDefaultVal As Single)
Declare Function NeuBiasGet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long) As Single
Declare Function NeuDeltaGet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long) As Single
Declare Function NeuValueGet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long) As Single
Declare Function NeuValueUpdatedGet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long) As Long
Declare Function NeuInputWeightGet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long, ByVal pLngInput As Long) As Single
Declare Function NeuInputWeightGetUpdated Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long, ByVal pLngInput As Long) As Long
Declare Function NeuInputsNumberGet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long) As Long
Declare Sub NeuBiasSet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long, ByVal pSngVal As Single)
Declare Sub NeuInputWeightSet Lib "AnaimoAI.dll" (ByVal pLngNeuron As Long, ByVal pLngInput As Long, ByVal pSngVal As Single)
Declare Sub NetOutputAdd Lib "AnaimoAI.dll" (ByVal pLngSrc As Long)
Declare Sub NetNeuronAdd Lib "AnaimoAI.dll" ()
Declare Function NetConnect Lib "AnaimoAI.dll" (ByVal pLngSrc As Long, ByVal pLngDst As Long) As Boolean
Declare Sub NetInputSet Lib "AnaimoAI.dll" (ByVal pLngInput As Long, ByVal pSngVal As Single)
Declare Function NetOutputGet Lib "AnaimoAI.dll" (ByVal pLngOutput As Long, ByVal pLngCyclesControl As Long) As Single
Declare Sub NetOutputSet Lib "AnaimoAI.dll" (ByVal pLngOutput As Long, ByVal pSngVal As Single)
Declare Sub NetLearningRateSet Lib "AnaimoAI.dll" (ByVal pSng As Single)
Declare Function NetLearn Lib "AnaimoAI.dll" (ByVal pLngCyclesControl As Long) As Boolean
Declare Sub NetSetDestroy Lib "AnaimoAI.dll" ()
'call NetSetStart after having NetSetRecord all records
Declare Function NetSetStart Lib "AnaimoAI.dll" (ByVal pSngDefaultVal As Single, ByVal pBolNetInitialize As Boolean) As Boolean
Declare Function NetSetRecord Lib "AnaimoAI.dll" () As Long
Declare Function NetSetLearnStart Lib "AnaimoAI.dll" (ByVal pSngThresholdForActive As Single, _
                                                     ByVal pSngDeviationPercentageTarget As Single, _
                                                     ByVal pLngCyclesControl As Long) As Boolean
Declare Function NetSetLearnContinue Lib "AnaimoAI.dll" (ByVal pLngRecordNumber As Long, ByVal pBolEstimateSuccess As Boolean, ByRef pBolThereIsNan As Boolean) As Single
Declare Function NetSetLearnEnd Lib "AnaimoAI.dll" () As Single
Declare Function NetSnapshotTake Lib "AnaimoAI.dll" () As Long
Declare Sub NetSnapshotGet Lib "AnaimoAI.dll" ()
Declare Function NetInputsMaxNumberGet Lib "AnaimoAI.dll" () As Long
Declare Function NetNeuronsAddedNumberGet Lib "AnaimoAI.dll" () As Long
Declare Function NetOutputsAddedNumberGet Lib "AnaimoAI.dll" () As Long
Declare Function NetThreadsMaxNumberGet Lib "AnaimoAI.dll" () As Long
Declare Sub NetThreadsMaxNumberSet Lib "AnaimoAI.dll" (ByVal pLng As Long)

Public Function SafeUbound(pArray As Variant, _
                           Optional ByVal pIntDim As Integer) As Long

  On Error GoTo Errores
  
  If pIntDim > 1 Then
    SafeUbound = UBound(pArray, pIntDim)
  Else
    SafeUbound = UBound(pArray)
  End If
  
Fin:
  Exit Function

Errores:
  SafeUbound = -1
  Resume Fin

End Function

'adds an item a tLngPieces list
Public Sub AddToLngList(pPieces() As Long, ByVal pLngItem As Long)
  ReDim Preserve pPieces(SafeUbound(pPieces) + 1) As Long
  pPieces(SafeUbound(pPieces)) = pLngItem
End Sub

Private Function MyNetCreate(ByVal pLngMaxNeurons As Long, _
                             ByVal pLngMaxInputs As Long, _
                             ByVal pLngMaxOutputs As Long) As Integer

  MyNetCreate = NetCreate(pLngMaxNeurons, pLngMaxInputs, pLngMaxOutputs)

  If MyNetCreate = NetCreate_LicenseExpiresInLessThan30Days Or MyNetCreate = NetCreate_NotLicensed Then
    If MyNetCreate = NetCreate_LicenseExpiresInLessThan30Days Then
      MsgBox "Neural Network license will expire in less than 30 days. Please request your new licensed version by providing to Anaimo this hardware id (you can paste it as has been copied to the clipboard): " & MyHardwareIdAndIntoClipBoard(), vbExclamation
    Else
      MsgBox "Neural Network license is not present or expired. Please request your licensed version by providing to Anaimo this hardware id (you can paste it as has been copied to the clipboard): " & MyHardwareIdAndIntoClipBoard(), vbCritical
    End If
  End If

End Function

Public Function CreateNeuralNetwork(ByVal pLngMaxNumberOfNeurons As Long, _
                                    ByVal pLngInputs As Long, _
                                    ByVal pLngOutputs As Long, _
                                    ByVal pBytConnectHow As Byte, _
                                    ByVal pIntMaxNumberOfLayers As Integer, _
                                    pIntNumberOfLayers As Integer, _
                                    pStrLayersBounds As String) As Boolean
  Const PROC_NAME = "CreateNeuralNetwork"
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim X As Long
  Dim l As Long
  Dim lStr As String
  Dim lLngCount As Long
  Dim lLngLastLayerNumberOfNeurons As Long
  Dim lIntLayerNum As Integer
  Dim lStrHID As String
  Dim lIntRes As Integer
  'to connect to later layers
  Dim lNeuronNumber() As Long
  Dim lLngPreviousLayers As Long
  Dim lLayerNeuronsStartIndex() As Long
  Dim lLayerNeuronsEndIndex() As Long
  
  'random connections
  Dim lLngSrc As Long
  Dim lLngDst As Long
  
  Const RANDOM_CONNECTIONS = 1000000
  
  #If vbDebug = 0 Then
    On Error GoTo Errores
  #End If
  
  'initialize
  pIntNumberOfLayers = 0
  pStrLayersBounds = vbNullString
  mLngFinalNumberOfTotalNeurons = 0
  mLngInputs = 0
  Call NetDestroy
        
  If pLngOutputs > pLngInputs Then
    MsgBox "Neural network cannot be created as inputs (" & pLngInputs & ") must be greater or equal than outputs (" & pLngOutputs & ").", vbCritical, App.Title
  ElseIf pLngInputs = pLngOutputs Then
    mLngInputs = pLngInputs
    pIntNumberOfLayers = pIntMaxNumberOfLayers
    'calculates number of neurons
    mLngFinalNumberOfTotalNeurons = pLngInputs * pIntMaxNumberOfLayers
    If mLngFinalNumberOfTotalNeurons >= pLngMaxNumberOfNeurons Then
      MsgBox "Maximum number of neurons reached.", vbCritical, App.Title
    Else
      lIntRes = MyNetCreate(mLngFinalNumberOfTotalNeurons, pLngInputs, pLngOutputs)
      If lIntRes <= NetCreate_LicenseExpiresInLessThan30Days Then
        pStrLayersBounds = "Layer 1: 0->" & (pLngInputs - 1) & vbCrLf
        For k = 1 To pIntMaxNumberOfLayers
          For i = 1 To pLngInputs
            Call NetNeuronAdd
          Next i
        Next k
        'Sets the outputs (on neurons already created)
        For i = 0 To pLngOutputs - 1
          Call NetOutputAdd(mLngFinalNumberOfTotalNeurons - pLngOutputs + i)
        Next i
        lLngCount = 0
        'neuron connections
        For k = 2 To pIntMaxNumberOfLayers
          pStrLayersBounds = pStrLayersBounds & "Layer " & k & ": " & pLngInputs * (k - 1) & "->" & pLngInputs * k - 1 & vbCrLf
          For i = pLngInputs * (k - 2) To pLngInputs * (k - 1) - 1
            For j = 0 To pLngInputs - 1
              'Debug.Print "Connecting " & i & "->" & pLngInputs * (k - 1) + j
              If IIf(NetConnect(i, pLngInputs * (k - 1) + j), 1, 0) = 0 Then
                MsgBox "Could not initialize the neural network (1).", vbCritical, App.Title
                lIntLayerNum = -1 'to exit
                Exit For
              End If
            Next j
            If lIntLayerNum = -1 Then Exit For
          Next i
        Next k
        pStrLayersBounds = pStrLayersBounds & "Outputs: " & mLngFinalNumberOfTotalNeurons - pLngOutputs & "->" & mLngFinalNumberOfTotalNeurons - 1
        'success?
        CreateNeuralNetwork = lIntLayerNum <> -1
      Else
        MsgBox "Could not initialize the neural network (4).Error: " & lIntRes, vbCritical, App.Title
      End If
    End If
  Else
    mLngInputs = pLngInputs
    X = (pLngInputs - pLngOutputs) / (pIntMaxNumberOfLayers - 1) 'substracts 1 for the outputs
    If X < 1 Then
      MsgBox "Neural network cannot be created as reduction in between layers (" & X & ") must be greater than 1.", vbCritical, App.Title
    Else
      'calculates number of neurons
      For k = pLngInputs To pLngOutputs + X Step -X
        pIntNumberOfLayers = pIntNumberOfLayers + 1
        mLngFinalNumberOfTotalNeurons = mLngFinalNumberOfTotalNeurons + k
        lLngLastLayerNumberOfNeurons = k
      Next k
      'pIntNumberOfLayers is now +1 because it includes the outputs
      If mLngFinalNumberOfTotalNeurons >= pLngMaxNumberOfNeurons Then
        MsgBox "Maximum number of neurons reached.", vbCritical, App.Title
      ElseIf lLngLastLayerNumberOfNeurons <= pLngOutputs Then
        MsgBox "Cannot add the neurons for the outputs as the number of neurons in the last layer is less or equal than the number of needed outputs.", vbCritical, App.Title
      Else
        'adds number of outputs
        mLngFinalNumberOfTotalNeurons = mLngFinalNumberOfTotalNeurons + pLngOutputs
        If mLngFinalNumberOfTotalNeurons >= pLngMaxNumberOfNeurons Then
          MsgBox "Total number of neurons (" & mLngFinalNumberOfTotalNeurons & ") exceeds the maximum number (" & pLngMaxNumberOfNeurons & ") of possible neurons.", vbCritical, App.Title
        Else
          lIntRes = MyNetCreate(mLngFinalNumberOfTotalNeurons, pLngInputs, pLngOutputs)
          If lIntRes <= NetCreate_LicenseExpiresInLessThan30Days Then
            pStrLayersBounds = "Layer 1: 0->" & (pLngInputs - 1) & vbCrLf
            'same FOR than above (*)
            For k = pLngInputs To pLngOutputs + X Step -X
              For i = 1 To k
                Call NetNeuronAdd
              Next i
            Next k
            
            'adds one more layer for the outputs
            pIntNumberOfLayers = pIntNumberOfLayers + 1
            
            'adds outputs' neurons
            For i = 1 To pLngOutputs
              Call NetNeuronAdd
            Next i
                                          
            'Sets the outputs (on neurons already created)
            For i = 0 To pLngOutputs - 1
              Call NetOutputAdd(mLngFinalNumberOfTotalNeurons - pLngOutputs + i)
            Next i
            
            'random connections
            If pBytConnectHow = 2 Then
              Randomize
              pStrLayersBounds = pStrLayersBounds & RANDOM_CONNECTIONS & " connections" & vbCrLf
              For i = 0 To RANDOM_CONNECTIONS
                Do
                  lLngSrc = (mLngFinalNumberOfTotalNeurons - pLngOutputs) * Rnd()
                  lLngDst = mLngFinalNumberOfTotalNeurons * Rnd()
                Loop Until lLngSrc <> lLngDst
                If IIf(NetConnect(lLngSrc, lLngDst), 1, 0) = 0 Then
                  MsgBox "Could not initialize the neural network (1).", vbCritical, App.Title
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
            'CONNECTIONS of every neuron to all later layers
            ElseIf pBytConnectHow = 1 Then
              lLngCount = 0
              lIntLayerNum = 1
              'neuron connections
              For k = pLngInputs To pLngOutputs + X Step -X
                If lLngCount + 2 * k - X - 1 < mLngFinalNumberOfTotalNeurons Then
                  pStrLayersBounds = pStrLayersBounds & "Layer " & lIntLayerNum + 1 & ": " & (lLngCount + k) & "->" & (lLngCount + 2 * k - X - 1) & vbCrLf
                  lIntLayerNum = lIntLayerNum + 1
                  AddToLngList lLayerNeuronsStartIndex, SafeUbound(lNeuronNumber) + 1 ' lLngCount, although incorrect, produced better results
                  For i = lLngCount To lLngCount + k - 1
                    For j = lLngCount + k To lLngCount + 2 * k - X - 1
                      If IIf(NetConnect(i, j), 1, 0) = 0 Then
                        MsgBox "Could not initialize the neural network (1).", vbCritical, App.Title
                        lIntLayerNum = -1 'to exit
                        Exit For
                      Else
                        AddToLngList lNeuronNumber, i 'when was here, although incorrect, produced better results
                      End If
                    Next j
                    If lIntLayerNum = -1 Then
                      Exit For
                    Else
                      AddToLngList lNeuronNumber, i
                    End If
                  Next i
                  AddToLngList lLayerNeuronsEndIndex, SafeUbound(lNeuronNumber) ' lLngCount + k - 1, although incorrect, produced better results
                  'connects all previous layers to output layer
                  For l = 0 To lLngPreviousLayers - 1
                    For i = lLayerNeuronsStartIndex(l) To lLayerNeuronsEndIndex(l)
                      For j = lLngCount + k To lLngCount + 2 * k - X - 1
                        If IIf(NetConnect(lNeuronNumber(i), j), 1, 0) = 0 Then
                          MsgBox "Could not initialize the neural network (1).", vbCritical, App.Title
                          lIntLayerNum = -1 'to exit
                          Exit For
                        End If
                      Next j
                      If lIntLayerNum = -1 Then Exit For
                    Next i
                  Next l
                  lLngPreviousLayers = lLngPreviousLayers + 1
                End If
                If lIntLayerNum = -1 Then Exit For
                lLngCount = lLngCount + k
              Next k
        
              'if no error
              If lIntLayerNum <> -1 Then
                pStrLayersBounds = pStrLayersBounds & "Outputs: " & mLngFinalNumberOfTotalNeurons - pLngOutputs & "->" & mLngFinalNumberOfTotalNeurons - 1
                'connects last layer to pLngOutputs
                For i = mLngFinalNumberOfTotalNeurons - pLngOutputs - lLngLastLayerNumberOfNeurons To mLngFinalNumberOfTotalNeurons - pLngOutputs - 1
                  For j = mLngFinalNumberOfTotalNeurons - pLngOutputs To mLngFinalNumberOfTotalNeurons - 1
                    If IIf(NetConnect(i, j), 1, 0) = 0 Then
                      MsgBox "Could not initialize the neural network (2).", vbCritical, App.Title
                      lIntLayerNum = -1 'to exit
                      Exit For
                    End If
                  Next j
                  If lIntLayerNum = -1 Then Exit For
                Next i
                'connects all previous layers to output layer
                For l = 0 To lLngPreviousLayers - 1
                  For i = lLayerNeuronsStartIndex(l) To lLayerNeuronsEndIndex(l)
                    For j = mLngFinalNumberOfTotalNeurons - pLngOutputs To mLngFinalNumberOfTotalNeurons - 1
                      If IIf(NetConnect(lNeuronNumber(i), j), 1, 0) = 0 Then
                        MsgBox "Could not initialize the neural network (2).", vbCritical, App.Title
                        lIntLayerNum = -1 'to exit
                        Exit For
                      End If
                    Next j
                    If lIntLayerNum = -1 Then Exit For
                  Next i
                Next l
              End If
            ' CONNECTIONS only with next layer
            Else
              lLngCount = 0
              lIntLayerNum = 1
              'neuron connections
              For k = pLngInputs To pLngOutputs + X Step -X
                If lLngCount + 2 * k - X - 1 < mLngFinalNumberOfTotalNeurons Then
                  pStrLayersBounds = pStrLayersBounds & "Layer " & lIntLayerNum + 1 & ": " & (lLngCount + k) & "->" & (lLngCount + 2 * k - X - 1) & vbCrLf
                  lIntLayerNum = lIntLayerNum + 1
                  For i = lLngCount To lLngCount + k - 1
                    For j = lLngCount + k To lLngCount + 2 * k - X - 1
                      If IIf(NetConnect(i, j), 1, 0) = 0 Then
                        MsgBox "Could not initialize the neural network (1).", vbCritical, App.Title
                        lIntLayerNum = -1 'to exit
                        Exit For
                      End If
                    Next j
                    If lIntLayerNum = -1 Then Exit For
                  Next i
                End If
                If lIntLayerNum = -1 Then Exit For
                lLngCount = lLngCount + k
              Next k
              'if no error
              If lIntLayerNum <> -1 Then
                pStrLayersBounds = pStrLayersBounds & "Outputs: " & mLngFinalNumberOfTotalNeurons - pLngOutputs & "->" & mLngFinalNumberOfTotalNeurons - 1
                'connects last layer to pLngOutputs
                For i = mLngFinalNumberOfTotalNeurons - pLngOutputs - lLngLastLayerNumberOfNeurons To mLngFinalNumberOfTotalNeurons - pLngOutputs - 1
                  For j = mLngFinalNumberOfTotalNeurons - pLngOutputs To mLngFinalNumberOfTotalNeurons - 1
                    If IIf(NetConnect(i, j), 1, 0) = 0 Then
                      MsgBox "Could not initialize the neural network (2).", vbCritical, App.Title
                      lIntLayerNum = -1 'to exit
                      Exit For
                    End If
                  Next j
                  If lIntLayerNum = -1 Then Exit For
                Next i
              End If
            End If
            CreateNeuralNetwork = lIntLayerNum <> -1
          Else
            MsgBox "Could not initialize the neural network. Possibly: out of memory. Error: " & lIntRes, vbCritical, App.Title
          End If
        End If
      End If
    End If
  End If
      
Fin:
  Exit Function
  
Errores:
  MsgBox Err.Number & "-" & Err.Description, vbCritical, App.Title
  Resume Fin
    
End Function

Public Function MyHardwareId() As String
  Dim i As Long
  Dim j As Long
  
  For i = 0 To 3
    For j = 0 To 3
      MyHardwareId = MyHardwareId & Hex$(HardwareId(i, j)) & ":"
    Next j
  Next i
  
  MyHardwareId = Left$(MyHardwareId, Len(MyHardwareId) - 1)

End Function

Public Function MyHardwareIdAndIntoClipBoard() As String
  Dim lStrHID As String

  lStrHID = MyHardwareId()
  
  'ignore errors as sometimes Clipboard might not be available
  On Error Resume Next

  Clipboard.Clear
  Clipboard.SetText lStrHID

  MyHardwareIdAndIntoClipBoard = lStrHID

End Function
