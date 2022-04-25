#pragma once

#ifdef AnaimoAI_EXPORTS
#define AnaimoAI_API __declspec(dllexport)
#else
#define AnaimoAI_API __declspec(dllimport)
#endif

#define NetCreate_Success 0
#define NetCreate_LicenseExpiresInLessThan30Days 1
#define NetCreate_NotLicensed 2
#define NetCreate_OutOfMemory 3
#define NetCreate_UnknownError 4

#define NetLearn_Success 0
#define NetLearn_NAN 1
#define NetLearn_ThreadsError 2
#define NetLearn_SetHasNoRecords 3

#define ACTIVATION_F_Sigmoid 0
#define ACTIVATION_F_ReLU 1 //although is simpler and faster to compute (but as it generates NAN, you have to repeat the process and finally is more expensive to compute), in forecasts, generates NAN in bias and does not forecast correctly
#define ACTIVATION_F_FastSigmoid 2

#ifdef _FROMLIB //Added in: Project properties, C/C++, All options, Preprocessor Definitions
extern int _cdecl HardwareId(unsigned int pIntRow, unsigned int pIntCol);
extern void _cdecl NetActivationFunctionSet(int pInt);
extern int _cdecl NetCreate(int pIntMaxNeurons, int pIntMaxInputs, int pIntMaxOutputs);
extern void _cdecl NetDestroy();
extern float _cdecl NetErrorGet(int pIntCyclesControl);
extern void _cdecl NetInitialize(float pSngDefaultVal);
extern float _cdecl NeuBiasGet(int pLngNeuron);
extern float _cdecl NeuDeltaGet(int pLngNeuron);
extern float _cdecl NeuValueGet(int pLngNeuron);
extern int _cdecl NeuValueUpdatedGet(int pLngNeuron);
extern float _cdecl NeuInputWeightGet(int pLngNeuron, int pLngInput);
extern int _cdecl NeuInputWeightUpdatedGet(int pLngNeuron, int pLngInput);
extern void _cdecl NeuInputWeightSet(int pLngNeuron, int pLngInput, float pSng);
extern int _cdecl NeuInputsNumberGet(int pLngNeuron);
extern void _cdecl NeuBiasSet(int pLngNeuron, float pSng);
extern void _cdecl NetOutputAdd(int pLngSrc);
extern void _cdecl NetNeuronAdd();
extern bool _cdecl NetConnect(int pLngSrc, int pLngDst);
extern bool _cdecl NetConnectConsecutive(int pLngSrc1, int pLngSrc2, int pLngDst);
extern float _cdecl NetDropOutGet();
extern void _cdecl NetDropOutSet(float pSng);
extern void _cdecl NetInputSet(int pLngInput, float pSng);
extern bool _cdecl NetLayersAnalyze();
extern float _cdecl NetMomentumGet();
extern void _cdecl NetMomentumSet(float pSng);
extern float _cdecl NetOutputGet(int pLngOutput, int pIntCyclesControl);
extern void _cdecl NetOutputSet(int pLngOutput, float pSng);
extern float _cdecl NetLearningRateGet();
extern void _cdecl NetLearningRateSet(float pSng);
extern int _cdecl NetLearn(int pIntCyclesControl);
extern int _cdecl NetModeGet();
extern void _cdecl NetModeSet(int pInt);
extern void _cdecl NetSetDestroy();
extern bool _cdecl NetSetStart(float pSngDefaultVal, bool pBolNetInitialize);
extern bool _cdecl NetSetPrepare(int pIntTotalNumberOfRecords);
extern int _cdecl NetSetRecord();
extern int _cdecl NetSetLearnStart(float pSngThresholdForActive, float pSngDeviationPercentageTarget, int pIntCyclesControl);
extern float _cdecl NetSetLearnContinue(int pIntRecordNumber, bool pBolEstimateSuccess, bool *pBolThereIsNan);
extern float _cdecl NetSetLearnEnd();
extern int _cdecl NetSnapshotTake();
extern bool _cdecl NetSnapshotGet();
extern int _cdecl NetThreadsMaxNumberGet();
extern void _cdecl NetThreadsMaxNumberSet(int pInt);
#else
#ifdef _MSC_VER
extern "C" AnaimoAI_API  int _stdcall HardwareId(unsigned int pIntRow, unsigned int pIntCol);
extern "C" AnaimoAI_API  void _stdcall NetActivationFunctionSet(int pInt);
extern "C" AnaimoAI_API  int _stdcall NetCreate(int pIntMaxNeurons, int pIntMaxInputs, int pIntMaxOutputs);
extern "C" AnaimoAI_API  void _stdcall NetDestroy();
extern "C" AnaimoAI_API  float _stdcall NetErrorGet(int pIntCyclesControl);
extern "C" AnaimoAI_API  void _stdcall NetInitialize(float pSngDefaultVal);
extern "C" AnaimoAI_API  float _stdcall NeuBiasGet(int pLngNeuron);
extern "C" AnaimoAI_API  float _stdcall NeuDeltaGet(int pLngNeuron);
extern "C" AnaimoAI_API  float _stdcall NeuValueGet(int pLngNeuron);
extern "C" AnaimoAI_API  int _stdcall NeuValueUpdatedGet(int pLngNeuron);
extern "C" AnaimoAI_API  float _stdcall NeuInputWeightGet(int pLngNeuron, int pLngInput);
extern "C" AnaimoAI_API  int _stdcall NeuInputWeightUpdatedGet(int pLngNeuron, int pLngInput);
extern "C" AnaimoAI_API  void _stdcall NeuInputWeightSet(int pLngNeuron, int pLngInput, float pSngVal);
extern "C" AnaimoAI_API  int _stdcall NeuInputsNumberGet(int pLngNeuron);
extern "C" AnaimoAI_API  void _stdcall NeuBiasSet(int pLngNeuron, float pSng);
extern "C" AnaimoAI_API  void _stdcall NetOutputAdd(int pLngSrc);
extern "C" AnaimoAI_API  void _stdcall NetNeuronAdd();
extern "C" AnaimoAI_API  bool _stdcall NetConnect(int pLngSrc, int pLngDst);
extern "C" AnaimoAI_API  bool _stdcall NetConnectConsecutive(int pLngSrc1, int pLngSrc2, int pLngDst);
extern "C" AnaimoAI_API  float _stdcall NetDropOutGet();
extern "C" AnaimoAI_API  void _stdcall NetDropOutSet(float pSng);
extern "C" AnaimoAI_API  void _stdcall NetInputSet(int pLngInput, float pSng);
extern "C" AnaimoAI_API  bool _stdcall NetLayersAnalyze();
extern "C" AnaimoAI_API  float _stdcall NetMomentumGet();
extern "C" AnaimoAI_API  void _stdcall NetMomentumSet(float pSng);
extern "C" AnaimoAI_API  float _stdcall NetOutputGet(int pLngOutput, int pIntCyclesControl);
extern "C" AnaimoAI_API  void _stdcall NetOutputSet(int pLngOutput, float pSng);
extern "C" AnaimoAI_API  float _stdcall NetLearningRateGet();
extern "C" AnaimoAI_API  void _stdcall NetLearningRateSet(float pSng);
extern "C" AnaimoAI_API  int _stdcall NetLearn(int pIntCyclesControl);
extern "C" AnaimoAI_API  int _stdcall NetModeGet();
extern "C" AnaimoAI_API  void _stdcall NetModeSet(int pInt);
extern "C" AnaimoAI_API  void _stdcall NetSetDestroy();
extern "C" AnaimoAI_API  bool _stdcall NetSetStart(float pSngDefaultVal, bool pBolNetInitialize);
extern "C" AnaimoAI_API  bool _stdcall NetSetPrepare(int pIntTotalNumberOfRecords);
extern "C" AnaimoAI_API  int _stdcall NetSetRecord();
extern "C" AnaimoAI_API  int _stdcall NetSetLearnStart(float pSngThresholdForActive, float pSngDeviationPercentageTarget, int pIntCyclesControl);
extern "C" AnaimoAI_API  float _stdcall NetSetLearnContinue(int pIntRecordNumber, bool pBolEstimateSuccess, bool* pBolThereIsNan);
extern "C" AnaimoAI_API  float _stdcall NetSetLearnEnd();
extern "C" AnaimoAI_API  int _stdcall NetSnapshotTake();
extern "C" AnaimoAI_API  bool _stdcall NetSnapshotGet();
extern "C" AnaimoAI_API  int _stdcall NetThreadsMaxNumberGet();
extern "C" AnaimoAI_API  void _stdcall NetThreadsMaxNumberSet(int pInt);
#else
extern int HardwareId(unsigned int pIntRow, unsigned int pIntCol);
extern void NetActivationFunctionSet(int pInt);
extern int NetCreate(int pIntMaxNeurons, int pIntMaxInputs, int pIntMaxOutputs);
extern void NetDestroy();
extern float NetErrorGet(int pIntCyclesControl);
extern void NetInitialize(float pSngDefaultVal);
extern float NeuBiasGet(int pLngNeuron);
extern float NeuDeltaGet(int pLngNeuron);
extern float NeuValueGet(int pLngNeuron);
extern int NeuValueUpdatedGet(int pLngNeuron);
extern float NeuInputWeightGet(int pLngNeuron, int pLngInput);
extern int NeuInputWeightUpdatedGet(int pLngNeuron, int pLngInput);
extern void NeuInputWeightSet(int pLngNeuron, int pLngInput, float pSng);
extern int NeuInputsNumberGet(int pLngNeuron);
extern void NeuBiasSet(int pLngNeuron, float pSng);
extern void NetOutputAdd(int pLngSrc);
extern void NetNeuronAdd();
extern bool NetConnect(int pLngSrc, int pLngDst);
extern bool NetConnectConsecutive(int pLngSrc1, int pLngSrc2, int pLngDst);
extern float NetDropOutGet();
extern void NetDropOutSet(float pSng);
extern void NetInputSet(int pLngInput, float pSng);
extern bool NetLayersAnalyze();
extern float NetMomentumGet();
extern void NetMomentumSet(float pSng);
extern float NetOutputGet(int pLngOutput, int pIntCyclesControl);
extern void NetOutputSet(int pLngOutput, float pSng);
extern float NetLearningRateGet();
extern void NetLearningRateSet(float pSng);
extern int NetLearn(int pIntCyclesControl);
extern int NetModeGet();
extern void NetModeSet(int pInt);
extern void NetSetDestroy();
extern bool NetSetStart(float pSngDefaultVal, bool pBolNetInitialize);
extern bool NetSetPrepare(int pIntTotalNumberOfRecords);
extern int NetSetRecord();
extern int NetSetLearnStart(float pSngThresholdForActive, float pSngDeviationPercentageTarget, int pIntCyclesControl);
extern float NetSetLearnContinue(int pIntRecordNumber, bool pBolEstimateSuccess, bool *pBolThereIsNan);
extern float NetSetLearnEnd();
extern int NetSnapshotTake();
extern bool NetSnapshotGet();
extern int NetThreadsMaxNumberGet();
extern void NetThreadsMaxNumberSet(int pInt);
#endif
#endif
