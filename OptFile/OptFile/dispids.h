// dispids.h
// dispatch ids

#pragma once

// properties
#define				DISPID_DetectorInfo			0x0100
#define				DISPID_InputInfo			0x0101
#define				DISPID_SlitInfo				0x0102
#define				DISPID_NumGratingScans		0x0103
#define				DISPID_RadianceAvailable	0x0104
#define				DISPID_CalibrationStandard	0x0105
#define				DISPID_CalibrationMeasurement	0x0106
#define				DISPID_OpticalTransferFunction	0x0107
#define				DISPID_TimeStamp			0x0108
#define				DISPID_ConfigFile			0x0109
#define				DISPID_ADInfoString			0x010a
#define				DISPID_Comment				0x010b
#define				DISPID_AllowNegativeValues	0x010c
#define				DISPID_ExtraValueTitles		0x010d
#define				DISPID_SignalUnits			0x010e
//#define				DISPID_ADGainFactor			0x010f

// methods
#define				DISPID_Clone				0x0120
#define				DISPID_Setup				0x0121
#define				DISPID_GetGratingScan		0x0122
#define				DISPID_InitNew				0x0123
#define				DISPID_ReadConfig			0x0124
#define				DISPID_SetExp				0x0125
#define				DISPID_AddValue				0x0126
#define				DISPID_SaveToFile			0x0127
#define				DISPID_LoadFromFile			0x0128
#define				DISPID_ClearScanData		0x0129
#define				DISPID_SetSlitInfo			0x012a
#define				DISPID_SelectCalibrationStandard	0x012b
#define				DISPID_ClearCalibrationStandard	0x012c
#define				DISPID_LoadOpticalTransferFile	0x012d
#define				DISPID_DisplayRadiance		0x012e
#define				DISPID_InitAcquisition		0x012f
#define				DISPID_ClearReadonly		0x0130
#define				DISPID_GetADChannel			0x0131
#define				DISPID_SaveToString			0x0132
#define				DISPID_LoadFromString		0x0133
#define				DISPID_FormOpticalTransferFile	0x0134
#define				DISPID_SetInputDeviceInfo	0x0135
#define				DISPID_ReadINI				0x0136
#define				DISPID_WriteINI				0x0137
#define				DISPID_GetExtraValues		0x0138
#define				DISPID_SetTimeStamp			0x0139
#define				DISPID_GetADGainFactor		0x013a

// events
#define				DISPID_RequestSlitWidth		0x0140
#define				DISPID_RequestGetInputPort	0x0141
#define				DISPID_RequestSetInputPort	0x0142
#define				DISPID_RequestGetOutputPort	0x0143
#define				DISPID_RequestSetOutputPort	0x0144
#define				DISPID_RequestParentWindow	0x0145

// GratingScanInfo
#define				DISPID_GratingScanInfo_NodeName		0x0180
#define				DISPID_GratingScanInfo_grating		0x0181
#define				DISPID_GratingScanInfo_Filter		0x0182
#define				DISPID_GratingScanInfo_dispersion	0x0183
#define				DISPID_GratingScanInfo_Wavelengths	0x0184
#define				DISPID_GratingScanInfo_Signals		0x0185
#define				DISPID_GratingScanInfo_amDirty		0x0186
//#define				DISPID_GratingScanInfo_ExtraValues	0x0187
#define				DISPID_GratingScanInfo_loadFromXML	0x0188
#define				DISPID_GratingScanInfo_saveToXML	0x0189
#define				DISPID_GratingScanInfo_addData		0x018a
#define				DISPID_GratingScanInfo_ClearData	0x018b
#define				DISPID_GratingScanInfo_clearDirty	0x018c
#define				DISPID_GratingScanInfo_findIndex	0x018d
#define				DISPID_GratingScanInfo_arraySize	0x018e
#define				DISPID_GratingScanInfo_getValue		0x018f
#define				DISPID_GratingScanInfo_setValue		0x0190
#define				DISPID_GratingScanInfo_setExtraValue	0x0191
#define				DISPID_GratingScanInfo_extraValues	0x0192
#define				DISPID_GratingScanInfo_haveExtraValue	0x0193
#define				DISPID_GratingScanInfo_ExtraValuesTitles	0x0194
#define				DISPID_GratingScanInfo_SignalUnits	0x0195
#define				DISPID_GratingScanInfo_ExtraValueWavelengths	0x0196
#define				DISPID_GratingScanInfo_ExtraValueComponent	0x0197
#define				DISPID_GratingScanInfo_Detector		0x0198

// calibration standard
#define				DISPID_CalibrationStandard_amLoaded		0x0220
#define				DISPID_CalibrationStandard_wavelengths	0x0221
#define				DISPID_CalibrationStandard_loadFromXML	0x0222
#define				DISPID_CalibrationStandard_saveToXML	0x0223
#define				DISPID_CalibrationStandard_fileName		0x0224
#define				DISPID_CalibrationStandard_calibration	0x0225
#define				DISPID_CalibrationStandard_loadExcelFile	0x0226
#define				DISPID_CalibrationStandard_checkExcelFile	0x0227
#define				DISPID_CalibrationStandard_expectedDistance	0x0228
#define				DISPID_CalibrationStandard_sourceDistance	0x0229

// container
#define			DISPID_Container_CreateObject		0x0280
#define			DISPID_Container_OnPropRequestEdit	0x0281
#define			DISPID_Container_OnPropChanged		0x0282
#define			DISPID_Container_OnSinkEvent		0x0283
#define			DISPID_Container_CloseObject		0x0284
#define			DISPID_Container_AmbientProperty	0x0285
#define			DISPID_Container_GetNamedObject		0x0286
#define			DISPID_Container_CloseAllObjects	0x0287
#define			DISPID_Container_TryMNemonic		0x0288

// extended object
#define			DISPID_Extended_DoCreateObject		0x0300
#define			DISPID_Extended_DoCloseObject		0x0301
#define			DISPID_Extended_TryMNemonic			0x0302
