#pragma once

class CMyObject;
class CDetectorInfo;
class CGratingScanInfo;
class CInputInfo;
class CSlitInfo;
class CCalibrationStandard;
class DataGainFactor;

class CMyOPTFile
{
public:
	CMyOPTFile(CMyObject * pMyObject);
	~CMyOPTFile(void);
	// persistence
	HRESULT				LoadFromStream(
							IStream		*	pStm);
	HRESULT				SaveToStream(
							IStream		*	pStm,
							BOOL			fClearDirty);
	BOOL				LoadFromFile();
	BOOL				SaveToFile();
	BOOL				LoadFromFile(
							LPCTSTR			szFileName);
	BOOL				SaveToFile(
							LPCTSTR			szFileName,
							BOOL			fClearDirty);
	BOOL				LoadFromString(
							LPCTSTR			szString);
	BOOL				SaveToString(
							LPTSTR		*	szString,
							BOOL			fClearDirty);
	BOOL				GetAmDirty();
	long				GetSaveSize();
	BOOL				InitNew();
	BOOL				GetAmLoaded();
	// properties and methods
	BOOL				GetDetectorInfo(
							IDispatch	**	ppdisp);
	BOOL				GetInputInfo(
							IDispatch	**	ppdisp);
	BOOL				GetSlitInfo(
							IDispatch	**	ppdisp);
	long				GetNumGratingScans();
	BOOL				GetRadianceAvailable();
	BOOL				GetIrradianceAvailable();
	BOOL				GetCalibrationStandard(
							LPTSTR			szCalibration,
							UINT			nBufferSize);
	void				SetCalibrationStandard(
							LPCTSTR			szCalibration);
	BOOL				GetCalibrationStandard(
							IDispatch	**	ppdisp);
	BOOL				GetCalibrationMeasurement();
	void				SetCalibrationMeasurement(
							BOOL			fCalibrationMeasurement);
	BOOL				GetADInfoString(
							LPTSTR		*	szInfoString);
	void				SetADInfoString(
							LPCTSTR			szInfoString);
	BOOL				Clone(
							IDispatch	**	ppdisp);
	BOOL				Setup(
							HWND			hwnd,
							LPCTSTR			Part);
	BOOL				GetGratingScan(
							long			segment,
							IDispatch	**	ppdisp);
	void				ClearCalibrationStandard();
	// comment string
	BOOL				GetComment(
							LPTSTR			szComment,
							UINT			nBufferSize);
	void				SetComment(
							LPCTSTR			szComment);
	// _clsIDataSet properties and methods
	BOOL				ReadDataFile(
							LPCTSTR			fileName);
	BOOL				WriteDataFile(
							LPCTSTR			fileName);
	BOOL				GetWavelengths(
							long		*	nArray,
							double		**	ppWaves);
	BOOL				GetSpectra(
							long		*	nArray,
							double		**	ppSpectra,
							BOOL			bandPassCorrected = FALSE);
	BOOL				AddValue(
							double			NM, 
							double			Signal);
	BOOL				ViewSetup(
							HWND			hwnd);
	long				GetArraySize();
	void				SetCurrentExp(
							LPCTSTR			filter,
							long			grating,
							long			detector);
	// read the configuration file
	BOOL				ReadConfig(
							LPCTSTR			ConfigFile);
	// clear the scan data
	void				ClearScanData();
	// 
	void				SetSlitInfo(
							LPCTSTR			SlitTitle,
							LPCTSTR			SlitName,
							double			slitWidth);
	// calculate radiance
	BOOL				CalculateRadiance(
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppSignal);
	// calculate Irradiance
	BOOL				CalculateIrradiance(
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppSignal,
							BOOL			fRadianceCalc = FALSE);
	// optical transfer function flag
	BOOL				GetOpticalTransferFunction();
	// load optical transfer function file
	BOOL				LoadOpticalTransferFile(
							LPCTSTR			szFileName);
	// get the time stamp
	BOOL				GetTimeStamp(
							LPTSTR			szTimeStamp,
							UINT			nBufferSize);
	// set time stamp
	void				SetTimeStamp();
	// initialize acquisition
	BOOL				InitAcquisition(
							HWND			hwndParent,
							BOOL			fCalibration);
	// get the configuration file
	BOOL				GetConfigFile(
							LPTSTR			szConfigFile,
							UINT			nBufferSize);
	// clear readonly
	void				ClearReadonly();
	// find AD channel given wavelength
	long				GetADChannel(
							double			wavelength);
	// file name
	BOOL				GetFileName(
							LPTSTR			szFileName,
							UINT			nBufferSize);
	void				SetFileName(
							LPCTSTR			szFileName);
	// single value
	double				GetDataValue(
							double			wavelength);
	void				SetDataValue(
							double			wavelength,
							double			signal);
	// form optical transfer file
	BOOL				FormOpticalTransfer(
							LPCTSTR			szFilePath);
	// allow negative values flag
	BOOL				GetAllowNegativeValues();
	void				SetAllowNegativeValues(
							BOOL			fAllowNegative);
	// extra scan value
	void				ExtraScanValue(
							LPCTSTR			szTitle,
							double			wavelength,
							double			extraValue);
	BOOL				GetExtraValues(
							LPCTSTR			szExtraValue,
							long		*	nValues,
							double		**	ppExtraValues);
	BOOL				GetExtraValueWaves(
							LPCTSTR			szExtraValue,
							long		*	nValues,
							double		**	ppWaves);
	BOOL				HaveExtraValues(); 
	BOOL				GetExtraValueTitles(
							LPTSTR		**	ppszTitles,
							ULONG		*	nStrings);
	BOOL				GetExtraValueTitle(
							long			index,
							LPTSTR			szTitle,
							UINT			nBufferSize);
	long				GetNumberExtraScanArrays();
	BOOL				GetExtraScanArray(
							long			index,
							long		*	nArray,
							double		**	ppwaves,
							double		**	ppvalues,
							LPTSTR			szTitle,
							UINT			nBufferSize);
	// signal units
	BOOL				GetSignalUnits(
							LPTSTR			szUnits,
							UINT			nBufferSize);
	void				SetSignalUnits(
							LPCTSTR			szUnits);
	// default title
	void				GetDefaultTitle(
							LPTSTR			szTitle,
							UINT			nBufferSize);
	void				SetDefaultTitle(
							LPCTSTR			szTitle);
	// single scan data
	BOOL				GetSingleGratingScan(
							long			grating,
							LPCTSTR			szFilter,
							long			detector,
							long		*	nArray,
							double		**	ppWaves,
							double		**	ppSignals);
	// export to CSV file
	BOOL				ExportToCSV(
							LPCTSTR			szFilePath);
	// determine detector and detector gain
	BOOL				DetermineDetector(
							double			wavelength,
							LPTSTR			szDetector,
							LPTSTR			szDetectorGain);
	// determine data gain factor
	double				DetermineDataGainFactor(
							double			wavelength);
	// minimum signal value
	double				GetMinSignalValue();
	// determine the DC signal offset
	double				DetermineDCSignalOffset(
							double			wavelength);
protected:
	BOOL				GetNMScanData(
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppSignal);
	BOOL				GetNMScanNormalizedData(			// scan data normalized by dispersion
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppSignal);
	BOOL				GetNMReferenceCalibration(
							double			startWave,
							double			endWave,
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppCal);
	BOOL				FormOpticalTransfer();
	BOOL				GetRefCalibration(
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppCal);
	BOOL				SingleScanNMData(
							long			iScan,
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppSignal);
	BOOL				GetCalibrationData(
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppSignals);
	BOOL				FindClosestValue(
							double			wavelength,
							long		*	gratingScan,
							long		*	index);
	// determine the additive factor for optical transfer calculation
	void				DetermineAdditiveFactor();
	// get scan data between given wavelengths
	BOOL				GetScanSection(
							long		nData,
							double	*	pWaves,
							double	*	pSignals,
							double		minWave,
							double		maxWave,
							long	*	nOut,
							double	**	ppWavesOut,
							double	**	ppSignalsOut);
	// smooth array
	void				SmoothArray(
							long		nArray,
							double	*	pArray);
	// lamp distance factor
	double				GetLampDistanceFactor();
	// number of extra values
	long				GetNumberExtraValues(
							LPCTSTR		szTitle);
	// determine the calibration gain factor
	double				DetermineCalibrationGainFactor(
							double		wavelength);
	// determine detector info
	BOOL				DetermineDetectorInfo(
							double		wavelength,
							DataGainFactor*	pDataGainFactor);
	// set optical transfer function flag on/off
	void				SetOpticalTransferOn(
							BOOL		fSetOpticalTransfer);
private:
	CMyObject		*	m_pMyObject;
	// grating scans
	long				m_nScans;
	CGratingScanInfo**	m_ppScans;
	// detector info
	CDetectorInfo	*	m_pDetectorInfo;
	// input info
	CInputInfo		*	m_pInputInfo;
	// slit info
	CSlitInfo		*	m_pSlitInfo;
	// calibration standard
	CCalibrationStandard*	m_pCalibrationStandard;
	// am loaded flag
	BOOL				m_fAmLoaded;
	// active scan index
	long				m_nActiveScan;
	// is this a calibration measurement?
	BOOL				m_fCalibrationMeasurement;
	// is this an optical transfer calculation?
	BOOL				m_fOpticalTransferFunction;
	// time stamp
	TCHAR				m_szTimeStamp[MAX_PATH];
	// configuration file name
	TCHAR				m_szConfigFile[MAX_PATH];
	// data file name
	TCHAR				m_szFilePath[MAX_PATH];
	// value to add to the raw data to ensure positive values for optical transfer function calculation
	double				m_addValue;
	// comment string
	TCHAR				m_szComment[MAX_PATH];
	// flag allow negative values
	BOOL				m_AllowNegativeValues;
	// extra value
	TCHAR				m_szExtraValue[MAX_PATH];
	// signal units
	TCHAR				m_szSignalUnits[MAX_PATH];
	// default title
	TCHAR				m_szDefaultTitle[MAX_PATH];
};
