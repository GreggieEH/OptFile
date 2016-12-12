#include "stdafx.h"
#include "MyOPTFile.h"
#include "MyObject.h"
#include "DetectorInfo.h"
#include "GratingScanInfo.h"
#include "InputInfo.h"
#include "SciXML.h"
#include "SlitInfo.h"
#include "PropPageDetectorInfo.h"
#include "PropPageInputInfo.h"
#include "PropPageSlitInfo.h"
#include "PropPageCalibrationStandard.h"
#include "PropPageGratingScanInfo.h"
#include "MyGuids.h"
#include "Server.h"
#include "IniFile.h"
#include "CalibrationStandard.h"			// calibration standard
#include "InterpolateValues.h"
#include "dispids.h"
#include "PropPageExtraData.h"
#include "DataGainFactor.h"
#include "ScriptDispatch.h"

CMyOPTFile::CMyOPTFile(CMyObject * pMyObject) :
	m_pMyObject(pMyObject),
	// grating scans
	m_nScans(0),
	m_ppScans(NULL),
	// detector info
	m_pDetectorInfo(NULL),
	// input info
	m_pInputInfo(NULL),
	// slit info
	m_pSlitInfo(NULL),
	// calibration standard
	m_pCalibrationStandard(NULL),
	// am loaded flag
	m_fAmLoaded(FALSE),
	// active scan index
	m_nActiveScan(-1),
	// is this a calibration measurement?
	m_fCalibrationMeasurement(FALSE),
	// is this an optical transfer calculation?
	m_fOpticalTransferFunction(FALSE),
	// value to add to the raw data to ensure positive values for optical transfer function calculation
	m_addValue(0.0),
	// flag allow negative values
	m_AllowNegativeValues(FALSE),
	// total run time
	m_fHaveTotalRunTime(FALSE)
{
	// time stamp
	this->m_szTimeStamp[0]	= '\0';
	// configuration file name
	this->m_szConfigFile[0]	= '\0';
	// data file path
	m_szFilePath[0]		= '\0';
	// comment string
	this->m_szComment[0] = '\0';
	// extra value title
	this->m_szExtraValue[0] = '\0';
	this->SetSignalUnits(L"V");
	this->SetDefaultTitle(L"Voltage");
	this->m_szTotalRunTime[0] = '\0';
}

CMyOPTFile::~CMyOPTFile(void)
{
	this->ClearScanData();
	Utils_DELETE_POINTER(this->m_pDetectorInfo);
	Utils_DELETE_POINTER(this->m_pInputInfo);
	Utils_DELETE_POINTER(this->m_pSlitInfo);
	Utils_DELETE_POINTER(this->m_pCalibrationStandard);
}

// persistence

HRESULT CMyOPTFile::LoadFromStream(
							IStream		*	pStm)
{
	HRESULT				hr;
	STATSTG				statStg;
	LPTSTR				szString		= NULL;
	DWORD				dwStreamSize;
	DWORD				dwNRead;

	hr = pStm->Stat(&statStg, STATFLAG_NONAME);
	if (SUCCEEDED(hr))
	{
		dwStreamSize	= (DWORD) statStg.cbSize.QuadPart;
		szString		= (LPTSTR) CoTaskMemAlloc(dwStreamSize + sizeof(TCHAR));
		ZeroMemory((PVOID) szString, dwStreamSize + sizeof(TCHAR));
		hr = pStm->Read((void*) szString, dwStreamSize, &dwNRead);
		if (SUCCEEDED(hr)) hr = this->LoadFromString(szString) ? S_OK : E_FAIL;
		CoTaskMemFree((LPVOID) szString);
	}
	return hr;
}

HRESULT CMyOPTFile::SaveToStream(
							IStream		*	pStm,
							BOOL			fClearDirty)
{
	HRESULT				hr;
	LPTSTR				szString		= NULL;
	size_t				slen;
	DWORD				dwNWritten;

	if (this->SaveToString(&szString, fClearDirty))
	{
		StringCbLength(szString, STRSAFE_MAX_CCH * sizeof(TCHAR), &slen);
		hr = pStm->Write((const void*) szString, slen, &dwNWritten);
		CoTaskMemFree((LPVOID) szString);
	}
	else hr = E_FAIL;
	return hr;
}

BOOL CMyOPTFile::LoadFromFile()
{
	return this->LoadFromFile(this->m_szFilePath);
}

BOOL CMyOPTFile::SaveToFile()
{
	// check if total run time is available
	HRESULT			hr;
	CLSID			clsid;
	IUnknown	*	pUnk;
	IDispatch	*	pdisp;
	DISPID			dispid;
	long			totalRunTime;
	long			hours;
	long			minutes;
	long			seconds;

	hr = CLSIDFromProgID(L"Sciencetech.ScanTimer.1", &clsid);
	if (SUCCEEDED(hr))
	{
		hr = GetActiveObject(clsid, NULL, &pUnk);
		if (SUCCEEDED(hr))
		{
			hr = pUnk->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
			if (SUCCEEDED(hr))
			{
				Utils_GetMemid(pdisp, L"TotalScanTime", &dispid);
				totalRunTime = Utils_GetLongProperty(pdisp, dispid);
				if (totalRunTime > 0 && totalRunTime < 1000000)
				{
					this->m_fHaveTotalRunTime = TRUE;
					if (totalRunTime > 3600)
					{
						hours = totalRunTime / 3600;
						minutes = (totalRunTime % 3600) / 60;
						StringCchPrintf(this->m_szTotalRunTime, MAX_PATH, L"%1d Hours %1d Minutes", hours, minutes);
					}
					else
					{
						minutes = totalRunTime / 60;
						seconds = totalRunTime % 60;
						StringCchPrintf(this->m_szTotalRunTime, MAX_PATH, L"%1d minutes %1d seconds", minutes, seconds);
					}
				}


				pdisp->Release();
			}
			pUnk->Release();
		}
	}


	return this->SaveToFile(this->m_szFilePath, TRUE);
}

BOOL CMyOPTFile::LoadFromFile(
							LPCTSTR			szFileName)
{
//	HRESULT			hr;
	HANDLE			hFile;
	BOOL			fSuccess		= FALSE;
	DWORD			fileSize;
	BYTE		*	szFileBuffer	= NULL;
	DWORD			dwNRead;
	DWORD			fileOffset = 0;

	hFile = CreateFile(szFileName, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE != hFile)
	{
		fileSize	= GetFileSize(hFile, NULL);
		szFileBuffer= (LPBYTE) CoTaskMemAlloc(fileSize + sizeof(TCHAR));
		ZeroMemory((PVOID) szFileBuffer, fileSize + sizeof(TCHAR));
		if (ReadFile(hFile, (LPVOID) szFileBuffer, fileSize, &dwNRead, NULL))
		{
			// endianness flag
			if (255 == szFileBuffer[0] && 254 == szFileBuffer[1])
			{
				fileOffset = 2;
			}
			if (IsTextUnicode((const void*)&(szFileBuffer[fileOffset]), fileSize, 0))
			{
				fSuccess = this->LoadFromString((LPCTSTR)&(szFileBuffer[fileOffset]));
			}
			else
			{
				// convert to unicode
				LPTSTR		szString;
				int			nChars = MultiByteToWideChar(CP_ACP, 0, (LPCSTR) szFileBuffer, -1, NULL, 0);
				szString = (LPTSTR)CoTaskMemAlloc((nChars + 1) * sizeof(TCHAR));
				MultiByteToWideChar(CP_ACP, 0, (LPCSTR)szFileBuffer, -1, szString, nChars + 1);
				fSuccess = this->LoadFromString(szString);
				CoTaskMemFree((LPVOID)szString);
			}
		}
		CoTaskMemFree((LPVOID) szFileBuffer);
		CloseHandle(hFile);
	}
	if (fSuccess)
	{
		StringCchCopy(this->m_szFilePath, MAX_PATH, szFileName);
		if (0 == lstrcmpi(PathFindExtension(szFileName), L".otd"))
		{
			this->SetDefaultTitle(L"Optical Transfer Data");
			this->SetSignalUnits(L"V/nm/(cm^2)");
		}
		else
		{
			this->SetDefaultTitle(L"Voltage");
			this->SetSignalUnits(L"V");
		}
	}
	return fSuccess;
}

BOOL CMyOPTFile::SaveToFile(
							LPCTSTR			szFileName,
							BOOL			fClearDirty)
{
	LPTSTR			szString	= NULL;
	HANDLE			hFile;
	BOOL			fSuccess	= FALSE;
	size_t			slen;
	DWORD			dwNWritten;
	DWORD			dwError;

	hFile = CreateFile(szFileName, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE != hFile)
	{
		if (this->SaveToString(&szString, fClearDirty))
		{
			StringCbLength(szString, STRSAFE_MAX_CCH * sizeof(TCHAR), &slen);
			fSuccess = WriteFile(hFile, (LPCVOID) szString, slen, &dwNWritten, NULL);
			CoTaskMemFree((LPVOID) szString);
		}
		CloseHandle(hFile);
	}
	else
	{
		dwError = GetLastError();
	}
	return fSuccess;
}

BOOL CMyOPTFile::LoadFromString(
							LPCTSTR			szString)
{
	CSciXML			xml;
	BOOL			fSuccess	= FALSE;
	TCHAR			szRoot[MAX_PATH];
//	long			nChild;
	long			i;
//	TCHAR			szNodeName[MAX_PATH];
	IDispatch	*	pdispXML;
	long			nArr;
	LPTSTR		*	arrScans;
	IDispatch	*	pdispRoot;					// root node
	IDispatch	*	pdispTimeStamp;				// time stamp node
	IDispatch	*	pdispComment;				// comment node
	IDispatch	*	pdispExtraValue;			// extra value node
	IDispatch	*	pdispConfigFile;			// config file path
	TCHAR			szUnits[MAX_PATH];
	IDispatch	*	pdispDefaultTitle;			// default title
	CInputInfo		inputInfo;
	IDispatch	*	pdisp;

	if (this->GetAmLoaded()) return FALSE;		// can't load this twice
	if (xml.loadFromString(szString))
	{
		// check the root node name
		if (xml.GetrootName(szRoot, MAX_PATH) && 0 == lstrcmpi(szRoot, ROOT_NODE_NAME))
		{
			if (xml.getMyObject(&pdispXML))
			{
				fSuccess = TRUE;
				// number of grating scans
				nArr			= 0;
				nArr			= xml.GetNGratingScans(&nArr, NULL);
				if (nArr > 0)
				{
					arrScans		= new LPTSTR [nArr];
					for (i=0; i<nArr; i++)
						arrScans[i]	= NULL;
					xml.GetNGratingScans(&nArr, arrScans);
					this->m_nScans	= nArr;
					if (this->m_nScans > 0)
					{
						this->m_ppScans		= new CGratingScanInfo * [this->m_nScans];
						for (i=0; i<this->m_nScans; i++)
						{
							this->m_ppScans[i] = new CGratingScanInfo;
							this->m_ppScans[i]->InitNew();
							this->m_ppScans[i]->SetnodeName(arrScans[i]);
							this->m_ppScans[i]->loadFromXML(pdispXML);
							CoTaskMemFree((LPVOID) arrScans[i]);
						}
						delete [] arrScans;
					}
					if (this->m_ppScans[0]->GetSignalUnits(szUnits, MAX_PATH))
					{
						this->SetSignalUnits(szUnits);
					}
				}
				this->m_pDetectorInfo	= new CDetectorInfo;
				this->m_pDetectorInfo->InitNew();
				this->m_pDetectorInfo->loadFromXML(pdispXML);

				if (this->m_pDetectorInfo->GetPulseIntegrator())
				{
					this->SetDefaultTitle(L"Signal Pulse");
					this->SetSignalUnits(L"V*sec");
				}
				else
				{
					this->SetDefaultTitle(L"Voltage");
					this->SetSignalUnits(L"V");
				}
				this->m_pInputInfo = new CScriptDispatch(L"InputInfo.jvs.txt");
				if (this->m_pInputInfo->GetScript(&pdisp))
				{
					inputInfo.doInit(pdisp);
					inputInfo.InitNew();
					inputInfo.loadFromXML(pdispXML);
					double		fov;
					if (inputInfo.getRadianceAvailable(&fov))
					{
						BOOL		stat = TRUE;
					}
					pdisp->Release();
				}

		//		this->m_pInputInfo		= new CInputInfo;
		//		this->m_pInputInfo->InitNew();
		//		this->m_pInputInfo->loadFromXML(pdispXML);
		//		double		fov;
		//		if (this->m_pInputInfo->getRadianceAvailable(&fov))
		//		{
		//			BOOL		stat = TRUE;
		//		}
				this->m_pSlitInfo		= new CSlitInfo;
				this->m_pSlitInfo->InitNew();
				this->m_pSlitInfo->loadFromXML(pdispXML);
	//			if (this->GetCalibrationMeasurement())
		//		{
					this->m_pCalibrationStandard	= new CCalibrationStandard;
					this->m_pCalibrationStandard->InitNew();
					if (this->m_pCalibrationStandard->loadFromXML(pdispXML))
					{
						this->SetCalibrationMeasurement(TRUE);
					}
					else
					{
			//			Utils_DELETE_POINTER(this->m_pCalibrationStandard);
						this->SetCalibrationMeasurement(FALSE);
					}
	//			}
				// time stamp
				if (xml.GetrootNode(&pdispRoot))
				{
					if (xml.findNamedChild(pdispRoot, TEXT("TimeStamp"), &pdispTimeStamp))
					{
						xml.getNodeValue(pdispTimeStamp, this->m_szTimeStamp, MAX_PATH);
						pdispTimeStamp->Release();
					}
					if (xml.findNamedChild(pdispRoot, L"Comment", &pdispComment))
					{
						xml.getNodeValue(pdispComment, this->m_szComment, MAX_PATH);
						pdispComment->Release();
					}
					if (xml.findNamedChild(pdispRoot, L"ExtraValue", &pdispExtraValue))
					{
						xml.getNodeValue(pdispExtraValue, this->m_szExtraValue, MAX_PATH);
						pdispExtraValue->Release();
					}
					if (xml.findNamedChild(pdispRoot, L"ConfigFile", &pdispConfigFile))
					{
						xml.getNodeValue(pdispConfigFile, this->m_szConfigFile, MAX_PATH);
						pdispConfigFile->Release();
					}
					if (xml.findNamedChild(pdispRoot, L"DefaultTitle", &pdispDefaultTitle))
					{
						xml.getNodeValue(pdispDefaultTitle, this->m_szDefaultTitle, MAX_PATH);
						pdispDefaultTitle->Release();
					}
					pdispRoot->Release();
				}
				pdispXML->Release();
			}
		}
	}
	if (fSuccess) this->m_fAmLoaded	= TRUE;
	return fSuccess;
}

BOOL CMyOPTFile::SaveToString(
							LPTSTR		*	szString,
							BOOL			fClearDirty)
{
	CSciXML			xml;
	BOOL			fSuccess	= FALSE;
	long			i;
	IDispatch	*	pdispXML;
	size_t			slen;
	IDispatch	*	pdispRoot;		// root node
	IDispatch	*	pdispTimeStamp;	// time stamp
	IDispatch	*	pdispComment;	// comment node
	IDispatch	*	pdispExtraValue;	// extra value node
	IDispatch	*	pdispConfigFile;	// config file node
	TCHAR			szDefaultTitle[MAX_PATH];
	IDispatch	*	pdispDefaultTitle;
	IDispatch	*	pdisp;
	CInputInfo		inputInfo;

	*szString	= NULL;
//	if (NULL == this->m_ppScans) return FALSE;
	// set the root name
	xml.SetrootName(ROOT_NODE_NAME);
	if (xml.initNew())
	{
		if (xml.getMyObject(&pdispXML))
		{
			// create the scan data nodes
			if (NULL != this->m_ppScans)
			{
				for (i=0; i<this->m_nScans; i++)
				{
					this->m_ppScans[i]->saveToXML(pdispXML);
				}
			}
			if (NULL != this->m_pDetectorInfo)
			{
				this->m_pDetectorInfo->saveToXML(pdispXML);
			}
			if (NULL != this->m_pInputInfo)
			{
		//		this->m_pInputInfo->saveToXML(pdispXML);
				if (this->m_pInputInfo->GetScript(&pdisp))
				{
					inputInfo.doInit(pdisp);
					inputInfo.saveToXML(pdispXML);
					pdisp->Release();
				}
			}
			if (NULL != this->m_pSlitInfo)
			{
				this->m_pSlitInfo->saveToXML(pdispXML);
			}
			if (this->GetCalibrationMeasurement() && NULL != this->m_pCalibrationStandard)
			{
				this->m_pCalibrationStandard->saveToXML(pdispXML);
			}
			// time stamp
			if (xml.GetrootNode(&pdispRoot))
			{
				// time stamp
				StringCchLength(this->m_szTimeStamp, MAX_PATH, &slen);
				if (slen > 0)
				{
					if (xml.createNode(TEXT("TimeStamp"), &pdispTimeStamp))
					{
						xml.setNodeValue(pdispTimeStamp, this->m_szTimeStamp);
						xml.appendChildNode(pdispRoot, pdispTimeStamp);
						pdispTimeStamp->Release();
					}
				}
				// comment string
				StringCchLength(this->m_szComment, MAX_PATH, &slen);
				if (slen > 0)
				{
					if (xml.createNode(L"Comment", &pdispComment))
					{
						xml.setNodeValue(pdispComment, this->m_szComment);
						xml.appendChildNode(pdispRoot, pdispComment);
						pdispComment->Release();
					}
				}
				// extra values
				StringCchLength(this->m_szExtraValue, MAX_PATH, &slen);
				if (slen > 0)
				{
					if (xml.createNode(L"ExtraValue", &pdispExtraValue))
					{
						xml.setNodeValue(pdispExtraValue, this->m_szExtraValue);
						xml.appendChildNode(pdispRoot, pdispExtraValue);
						pdispExtraValue->Release();
					}
				}
				// config file
				StringCchLength(this->m_szConfigFile, MAX_PATH, &slen);
				if (slen > 0)
				{
					if (xml.createNode(L"ConfigFile", &pdispConfigFile))
					{
						xml.setNodeValue(pdispConfigFile, this->m_szConfigFile);
						xml.appendChildNode(pdispRoot, pdispConfigFile);
						pdispConfigFile->Release();
					}
				}
				// default title
				this->GetDefaultTitle(szDefaultTitle, MAX_PATH);
				if (xml.createNode(L"DefaultTitle", &pdispDefaultTitle))
				{
					xml.setNodeValue(pdispDefaultTitle, szDefaultTitle);
					xml.appendChildNode(pdispRoot, pdispDefaultTitle);
					pdispDefaultTitle->Release();
				}
				if (this->m_fHaveTotalRunTime)
				{
					if (xml.createNode(L"TotalRunTime", &pdisp))
					{
						xml.setNodeValue(pdisp, this->m_szTotalRunTime);
						xml.appendChildNode(pdispRoot, pdisp);
						pdisp->Release();
					}
				}
				pdispRoot->Release();
			}
			// write to string
			fSuccess = xml.saveToString(szString);
			if (fSuccess && fClearDirty)
			{
				for (i=0; i<this->m_nScans; i++)
					this->m_ppScans[i]->clearDirty();
				if (NULL != this->m_pDetectorInfo) this->m_pDetectorInfo->clearDirty();
				if (NULL != this->m_pInputInfo)
				{
					inputInfo.clearDirty();
				}
				if (NULL != this->m_pSlitInfo) this->m_pSlitInfo->clearDirty();
			}
			pdispXML->Release();
		}
	}
	return fSuccess;
}

BOOL CMyOPTFile::GetAmDirty()
{
	BOOL			fAmDirty	= FALSE;
	long			i			= 0;
	IDispatch	*	pdisp;

	// check the grating scans
	while (i < this->m_nScans && !fAmDirty)
	{
		fAmDirty = this->m_ppScans[i]->GetamDirty();
		if (!fAmDirty) i++;
	}
	if (fAmDirty) return TRUE;
	if (NULL != this->m_pDetectorInfo) if (this->m_pDetectorInfo->GetamDirty()) return TRUE;

//	if (NULL != this->m_pInputInfo) if (this->m_pInputInfo->GetamDirty()) return TRUE;
	if (NULL != this->m_pInputInfo)
	{
		if (this->m_pInputInfo->GetScript(&pdisp))
		{
			CInputInfo		inputInfo;
			inputInfo.doInit(pdisp);
			fAmDirty = inputInfo.GetamDirty();
			pdisp->Release();
		}
		if (fAmDirty) return TRUE;
	}
	if (NULL != this->m_pSlitInfo) if (this->m_pSlitInfo->GetamDirty()) return TRUE;
	return FALSE;
}

long CMyOPTFile::GetSaveSize()
{
	LPTSTR			szString	= NULL;
	size_t			slen		= 0;

	if (this->SaveToString(&szString, FALSE))
	{
		StringCchLength(szString, STRSAFE_MAX_CCH, &slen);
		CoTaskMemFree((LPVOID) szString);
	}
	return (long) slen;
}

BOOL CMyOPTFile::InitNew()
{
	if (this->GetAmLoaded()) return FALSE;
	this->m_pDetectorInfo	= new CDetectorInfo;
	this->m_pDetectorInfo->InitNew();
	this->m_pInputInfo = new CScriptDispatch(L"InputInfo.jvs.txt");
	IDispatch	*	pdisp;
	if (this->m_pInputInfo->GetScript(&pdisp))
	{
		CInputInfo		inputInfo;
		inputInfo.doInit(pdisp);
		inputInfo.InitNew();
		pdisp->Release();
	}
//		= new CInputInfo;
//	this->m_pInputInfo->InitNew();
	this->m_pSlitInfo		= new CSlitInfo;
	this->m_pSlitInfo->InitNew();
	this->m_pCalibrationStandard	= new CCalibrationStandard;
	this->m_pCalibrationStandard->InitNew();
//	this->m_fAmLoaded	= TRUE;
	return TRUE;
}

BOOL CMyOPTFile::GetAmLoaded()
{
	return this->m_fAmLoaded;
}

// properties and methods
BOOL CMyOPTFile::GetDetectorInfo(
							IDispatch	**	ppdisp)
{
	if (NULL != this->m_pDetectorInfo)
	{
		return this->m_pDetectorInfo->getMyObject(ppdisp);
	}
	else
	{
		*ppdisp		= NULL;
		return FALSE;
	}
}

BOOL CMyOPTFile::GetInputInfo(
							IDispatch	**	ppdisp)
{
	if (NULL != this->m_pInputInfo)
	{
		return this->m_pInputInfo->GetScript(ppdisp);
	}
	else
	{
		*ppdisp	= NULL;
		return FALSE;
	}
}

BOOL CMyOPTFile::GetSlitInfo(
							IDispatch	**	ppdisp)
{
	if (NULL != this->m_pSlitInfo)
	{
		if (!this->m_pSlitInfo->AmInit())
		{
			this->m_pSlitInfo->InitNew();
		}
		return this->m_pSlitInfo->getMyObject(ppdisp);
	}
	else
	{
		*ppdisp	= NULL;
		return FALSE;
	}
}

long CMyOPTFile::GetNumGratingScans()
{
	return this->m_nScans;
}

BOOL CMyOPTFile::GetRadianceAvailable()
{
	IDispatch	*	pdisp;
	BOOL			radianceAvailable = FALSE;
	CInputInfo		inputInfo;
	double			fov;

	if (this->GetInputInfo(&pdisp))
	{
		inputInfo.doInit(pdisp);
		radianceAvailable = inputInfo.getRadianceAvailable(&fov);
		pdisp->Release();
	}
	return radianceAvailable;
/*
	if (NULL != this->m_pInputInfo)
	{
		double				FOV;
		return this->m_pInputInfo->getRadianceAvailable(&FOV);
	}
	else
		return FALSE;
*/
}

BOOL CMyOPTFile::GetIrradianceAvailable()
{
	return TRUE;
}

BOOL CMyOPTFile::GetCalibrationStandard(
							IDispatch	**	ppdisp)
{
	if (NULL != this->m_pCalibrationStandard)
	{
		return this->m_pCalibrationStandard->getMyObject(ppdisp);
	}
	else
	{
		*ppdisp		= NULL;
		return FALSE;
	}
}

BOOL CMyOPTFile::GetCalibrationStandard(
							LPTSTR			szCalibration,
							UINT			nBufferSize)
{
	if (NULL != this->m_pCalibrationStandard)
	{
		return this->m_pCalibrationStandard->GetfileName(szCalibration, nBufferSize);
	}
	else
	{
		szCalibration[0]	= '\0';
		return FALSE;
	}
}

void CMyOPTFile::SetCalibrationStandard(
							LPCTSTR			szCalibration)
{
	if (NULL == this->m_pCalibrationStandard)
	{
		this->m_pCalibrationStandard	= new CCalibrationStandard();
		this->m_pCalibrationStandard->InitNew();
	}
	if (NULL != this->m_pCalibrationStandard)
	{
		this->m_pCalibrationStandard->SetfileName(szCalibration);
		this->SetCalibrationMeasurement(TRUE);
	}
}

BOOL CMyOPTFile::GetCalibrationMeasurement()
{
	BOOL			fSuccess = FALSE;
	// make sure that this is a calibration measurement
	if (this->m_fCalibrationMeasurement)
	{
		if (NULL != this->m_pCalibrationStandard)
		{
			if (!this->m_pCalibrationStandard->GetamLoaded())
				this->m_pCalibrationStandard->SetamLoaded(TRUE);
			fSuccess = this->m_pCalibrationStandard->GetamLoaded();
		}
		if (!fSuccess) this->m_fCalibrationMeasurement = FALSE;
	}

	return this->m_fCalibrationMeasurement;
}

void CMyOPTFile::SetCalibrationMeasurement(
							BOOL			fCalibrationMeasurement)
{
	this->m_fCalibrationMeasurement = fCalibrationMeasurement;
}

BOOL CMyOPTFile::GetADInfoString(
	LPTSTR		*	szInfoString)
{
	return this->m_pDetectorInfo->GetADInfoString(szInfoString);
}

void CMyOPTFile::SetADInfoString(
	LPCTSTR			szInfoString)
{
	this->m_pDetectorInfo->SetADInfoString(szInfoString);
	if (this->m_pDetectorInfo->GetPulseIntegrator())
	{
		this->SetDefaultTitle(L"Signal Pulse");
		this->SetSignalUnits(L"V*sec");
	}
	else
	{
		this->SetDefaultTitle(L"Voltage");
		this->SetSignalUnits(L"V");
	}
}

BOOL CMyOPTFile::Clone(
							IDispatch	**	ppdisp)
{
	HRESULT				hr;
	IStream			*	pStm;
	IPersistStreamInit*	pPersistStream;
	BOOL				fSuccess	= FALSE;
	IUnknown		*	punk;
	LARGE_INTEGER		dlibMove;
	*ppdisp		= NULL;

	// create temp stream
	hr = CreateStreamOnHGlobal(NULL, TRUE, &pStm);
	if (SUCCEEDED(hr))
	{
		// write our object to a stream
		hr = this->m_pMyObject->QueryInterface(IID_IPersistStreamInit, (LPVOID*) &pPersistStream);
		if (SUCCEEDED(hr))
		{
			hr = pPersistStream->Save(pStm, FALSE);
			pPersistStream->Release();
		}
		// create clone
		hr = CoCreateInstance(MY_CLSID, NULL, CLSCTX_INPROC_SERVER, 
			IID_IUnknown, (LPVOID*) &punk);
		if (SUCCEEDED(hr))
		{
			// load cloned object from stream
			hr = punk->QueryInterface(IID_IPersistStreamInit, (LPVOID*) &pPersistStream);
			if (SUCCEEDED(hr))
			{
				// reset the stream position to start
				dlibMove.QuadPart	= 0;
				pStm->Seek(dlibMove, STREAM_SEEK_SET, NULL);
				hr = pPersistStream->Load(pStm);
				pPersistStream->Release();
			}
			if (SUCCEEDED(hr))
			{
				hr = punk->QueryInterface(IID_IDispatch, (LPVOID*) ppdisp);
				fSuccess = SUCCEEDED(hr);
				if (fSuccess)
				{
					Utils_InvokeMethod(*ppdisp, DISPID_ClearReadonly, NULL, 0, NULL);
				}
			}
			punk->Release();
		}
		pStm->Release();
	}
	return fSuccess;
}

BOOL CMyOPTFile::Setup(
							HWND			hwnd,
							LPCTSTR			Part)
{
	PROPSHEETHEADER			psh;
	UINT					nPages;
	HPROPSHEETPAGE		*	aPages;
	CPropPageDetectorInfo	propPageDetector(this->m_pMyObject);
	CPropPageInputInfo		propPageInput(this->m_pMyObject);
	CPropPageSlitInfo		propPageSlitInfo(this->m_pMyObject);
	CPropPageCalibrationStandard	propPageCalStandard(this->m_pMyObject);
	CPropPageGratingScanInfo	propPageGratingScan(this->m_pMyObject);
	CPropPageExtraData		propPageExtraData(this->m_pMyObject);
	BOOL					fSuccess	= FALSE;

	if (0 == lstrcmpi(Part, TEXT("Setup")))
	{
		if (this->m_nScans > 0)
		{
			BOOL			fExtraValues = this->HaveExtraValues();
			if (this->m_fCalibrationMeasurement)
			{
				nPages = fExtraValues ? 6 : 5;
				aPages = new HPROPSHEETPAGE[nPages];
				aPages[0] = propPageInput.doCreatePropPage();
				aPages[1] = propPageDetector.doCreatePropPage();
				aPages[2] = propPageSlitInfo.doCreatePropPage();
				aPages[3] = propPageCalStandard.doCreatePropPage();
				aPages[4] = propPageGratingScan.doCreatePropPage();
				if (fExtraValues)
					aPages[5] = propPageExtraData.doCreatePropPage();
			}
			else
			{
				nPages = fExtraValues ? 5 : 4;
				aPages = new HPROPSHEETPAGE[nPages];
				aPages[0] = propPageInput.doCreatePropPage();
				aPages[1] = propPageDetector.doCreatePropPage();
				aPages[2] = propPageSlitInfo.doCreatePropPage();
				aPages[3] = propPageGratingScan.doCreatePropPage();
				if (fExtraValues)
					aPages[4] = propPageExtraData.doCreatePropPage();
			}
		}
		else
		{
			if (this->m_fCalibrationMeasurement)
			{
				nPages = 4;
				aPages = new HPROPSHEETPAGE[nPages];
				aPages[0] = propPageInput.doCreatePropPage();
				aPages[1] = propPageDetector.doCreatePropPage();
				aPages[2] = propPageSlitInfo.doCreatePropPage();			// don't display slit info if no data
				aPages[2] = propPageCalStandard.doCreatePropPage();
			}
			else
			{
				nPages = 3;
				aPages = new HPROPSHEETPAGE[nPages];
				aPages[0] = propPageInput.doCreatePropPage();
				aPages[1] = propPageDetector.doCreatePropPage();
				aPages[2] = propPageSlitInfo.doCreatePropPage();			// don't display slit info if no data
			}
		}
	}
	else if (0 == lstrcmpi(Part, TEXT("Input")))
	{
		nPages		= 1;
		aPages		= new HPROPSHEETPAGE [nPages];
		aPages[0]	= propPageInput.doCreatePropPage();
	}
	else if (0 == lstrcmpi(Part, TEXT("Detector")))
	{
		nPages		= 1;
		aPages		= new HPROPSHEETPAGE [nPages];
		aPages[0]	= propPageDetector.doCreatePropPage();
	}
	else if (0 == lstrcmpi(Part, TEXT("SlitInfo")))
	{
		nPages		= 1;
		aPages		= new HPROPSHEETPAGE [nPages];
		aPages[0]	= propPageSlitInfo.doCreatePropPage();
	}
	else if (0 == lstrcmpi(Part, TEXT("Spectral Data")))
	{
		nPages		= 1;
		aPages		= new HPROPSHEETPAGE [nPages];
		aPages[0]	= propPageGratingScan.doCreatePropPage();
	}
	else if (0 == lstrcmpi(Part, TEXT("Calibration Standard")))
	{
		nPages		= 1;
		aPages		= new HPROPSHEETPAGE [nPages];
		aPages[0]	= propPageCalStandard.doCreatePropPage();
	}
	else if (0 == lstrcmpi(Part, TEXT("Slit And Calibration")))
	{
		nPages		= 2;
		aPages		= new HPROPSHEETPAGE [nPages];
		aPages[0]	= propPageSlitInfo.doCreatePropPage();
		aPages[1]	= propPageCalStandard.doCreatePropPage();
	}
	else if (0 == lstrcmpi(Part, TEXT("Input And Detector")))
	{
		nPages = 2;
		aPages = new HPROPSHEETPAGE[nPages];
		aPages[0] = propPageInput.doCreatePropPage();
		aPages[1] = propPageDetector.doCreatePropPage();
	}
	else
	{
		return FALSE;
	}
	ZeroMemory((PVOID) &psh, sizeof(PROPSHEETHEADER));
	psh.dwSize		= sizeof(PROPSHEETHEADER);
	psh.dwFlags		= PSH_DEFAULT;
	psh.hwndParent	= hwnd;
	psh.hInstance	= GetServer()->GetInstance();
	psh.nPages		= nPages;
	psh.phpage		= aPages;
	fSuccess = PropertySheet(&psh) > 0;
	delete [] aPages;
	return fSuccess;
}

BOOL CMyOPTFile::GetGratingScan(
							long			segment,
							IDispatch	**	ppdisp)
{
	if (segment >= 0 && segment < this->m_nScans)
	{
		return this->m_ppScans[segment]->getMyObject(ppdisp);
	}
	else
	{
		*ppdisp		= NULL;
		return FALSE;
	}
}

void CMyOPTFile::ClearCalibrationStandard()
{
	Utils_DELETE_POINTER(this->m_pCalibrationStandard);
	this->SetCalibrationMeasurement(FALSE);
}

// _clsIDataSet properties and methods
BOOL CMyOPTFile::ReadDataFile(
							LPCTSTR			fileName)
{
	return this->LoadFromFile(fileName);
}

BOOL CMyOPTFile::WriteDataFile(
							LPCTSTR			fileName)
{
	return this->SaveToFile(fileName, TRUE);
}

BOOL CMyOPTFile::GetWavelengths(
							long		*	nArray,
							double		**	ppWaves)
{
	long		arraySize		= this->GetArraySize();
	long		iScan;
	long		arrOffset		= 0;
	long		i;
	long		nValues;
	double	*	pWaves;

	if (0 == arraySize)
	{
		*nArray		= 0;
		*ppWaves	= NULL;
		return FALSE;
	}
	*nArray		= arraySize;
	*ppWaves	= new double [arraySize];
	for (iScan=0; iScan<this->m_nScans; iScan++)
	{
		if (this->m_ppScans[iScan]->Getwavelengths(&nValues, &pWaves))
		{
			for (i=0; i<nValues; i++)
			{
				(*ppWaves)[arrOffset + i] = pWaves[i];
			}
			// increment the array offset
			arrOffset += nValues;
			delete [] pWaves;
		}
	}
	return TRUE;
}

BOOL CMyOPTFile::GetSpectra(
							long		*	nArray,
							double		**	ppSpectra,
							BOOL			bandPassCorrected /*= FALSE*/)
{
	long		arraySize		= this->GetArraySize();
	long		iScan;
	long		arrOffset		= 0;
	long		i;
	long		nValues;
	double	*	pSpectra;
	double		outputSlitWidth;		//	= this->m_pSlitInfo->GetOutputSlitWidth();
	double		bandPass;
//	double		lampDistanceFactor = this->GetLampDistanceFactor();
	BOOL		allowNegative = this->GetAllowNegativeValues();
	double		minValue = 2.0 / 65535.0;

	if (0 == arraySize)
	{
		*nArray		= 0;
		*ppSpectra	= NULL;
		return FALSE;
	}
	this->m_pSlitInfo->GetOutputSlitWidth(&outputSlitWidth);
	*nArray		= arraySize;
	*ppSpectra	= new double [arraySize];
	for (iScan=0; iScan<this->m_nScans; iScan++)
	{
		bandPass = outputSlitWidth * this->m_ppScans[iScan]->Getdispersion();
		if (0.0 == bandPass) bandPass = 1.0;
		if (this->m_ppScans[iScan]->Getsignals(&nValues, &pSpectra))
		{
			for (i=0; i<nValues; i++)
			{
				if (!allowNegative && pSpectra[i] < minValue)
					pSpectra[i] = minValue;

				(*ppSpectra)[arrOffset + i] = pSpectra[i];
				if (bandPassCorrected)
				{
					// correct for bandpass and lamp distance factor
					(*ppSpectra)[arrOffset + i] /= bandPass;
				}
			}
			// increment the array offset
			arrOffset += nValues;
			delete []pSpectra;
		}
	}
	return TRUE;
}

BOOL CMyOPTFile::AddValue(
							double			NM, 
							double			Signal)
{
	if (this->m_nActiveScan >= 0 && this->m_nActiveScan < this->m_nScans)
	{
		this->m_ppScans[this->m_nActiveScan]->addData(NM, Signal);
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

BOOL CMyOPTFile::ViewSetup(
							HWND			hwnd)
{
	return this->Setup(hwnd, TEXT("Setup"));
}

long CMyOPTFile::GetArraySize()
{
	long			arraySize	= 0;
	long			iScan;
	for (iScan=0; iScan<this->m_nScans; iScan++)
		arraySize += this->m_ppScans[iScan]->GetarraySize();
	return arraySize;
}

void CMyOPTFile::SetCurrentExp(
							LPCTSTR			filter, 
							long			grating,
							long			detector)
{
	BOOL			fDone	= FALSE;
	long			iScan	= 0;
	TCHAR			szNodeName[MAX_PATH];

//	TCHAR			szMessage[MAX_PATH];
//	StringCchPrintf(szMessage, MAX_PATH, L"SetCurrentExp for filter %s grating %1d detector %1d", filter, grating, detector);
//	MessageBox(NULL, szMessage, L"SetCurrentExp", MB_OK);

	while (iScan < this->m_nScans && !fDone)
	{
		fDone = this->m_ppScans[iScan]->CheckGratingAndFilter(
			grating, filter, detector);
		if (fDone)
		{
			this->m_nActiveScan	= iScan;
		}
		if (!fDone) iScan++;
	}
	if (!fDone)
	{

		// need to add a scan
		long				nScans	= this->m_nScans + 1;
		CGratingScanInfo**	ppScans	= new CGratingScanInfo* [nScans];
		for (iScan=0; iScan<this->m_nScans; iScan++)
			ppScans[iScan] = this->m_ppScans[iScan];
		ppScans[iScan]		= new CGratingScanInfo;
		ppScans[iScan]->InitNew();
		ppScans[iScan]->SetSignalUnits(this->m_szSignalUnits);
		// set the node name
		StringCchPrintf(szNodeName, MAX_PATH, TEXT("SpectraData_%1d"), iScan+1);
		ppScans[iScan]->SetnodeName(szNodeName);
		ppScans[iScan]->Setfilter(filter);
		ppScans[iScan]->Setgrating(grating);
		ppScans[iScan]->SetDetector(detector);
		// request the dispersion
		ppScans[iScan]->Setdispersion(this->m_pMyObject->FirerequestDispersion(grating));
		if (NULL != this->m_ppScans) delete [] this->m_ppScans;
		this->m_ppScans	= ppScans;
		this->m_nScans	= nScans;
		this->m_nActiveScan	= iScan;

//		StringCchPrintf(szMessage, MAX_PATH, L"Added scan for filter %s grating %1d detector %1d", filter, grating, detector);
//		MessageBox(NULL, szMessage, L"OptFile", MB_OK);
	}
}

// read the configuration file
BOOL CMyOPTFile::ReadConfig(
							LPCTSTR			ConfigFile)
{
	BOOL			fSuccess	= FALSE;
	CIniFile		iniFile;
	long			lval;
	TCHAR			szString[16];

	if (iniFile.Init(ConfigFile))
	{
		iniFile.ReadIntValue(TEXT("OPTMeasure"), TEXT("Available"), 0, &lval);
		if (0 == lval) return FALSE;
		fSuccess = TRUE;
		this->InitNew();
		// detector info
		if (iniFile.ReadStringValue(TEXT("OPTDetector"), TEXT("Model"), TEXT("PMT"), szString, 16))
		{
			StrTrim(szString, TEXT(" "));
			this->m_pDetectorInfo->SetdetectorModel(szString);
		}
	}
	if (fSuccess)
	{
		// store the configuration file name
		StringCchCopy(this->m_szConfigFile, MAX_PATH, ConfigFile);
	}
	return fSuccess;
}

// clear the scan data
void CMyOPTFile::ClearScanData()
{
	long			nScans;
	long			i;

	if (NULL != this->m_ppScans)
	{
		nScans	= this->m_nScans;
		for (i=0; i<nScans; i++)
		{
			Utils_DELETE_POINTER(this->m_ppScans[i]);
		}
		delete [] this->m_ppScans;
		this->m_ppScans	= NULL;
	}
	this->m_nScans	= 0;
}

void CMyOPTFile::SetSlitInfo(
							LPCTSTR			SlitTitle,
							LPCTSTR			SlitName,
							double			slitWidth)
{
	if (NULL == this->m_pSlitInfo) return;
	if (this->m_pSlitInfo->haveSlit(SlitTitle))
	{
		this->m_pSlitInfo->setSlitName(SlitTitle, SlitName);
		this->m_pSlitInfo->setSlitWidth(SlitTitle, slitWidth);
	}
	else
	{
		this->m_pSlitInfo->addSlit(SlitTitle, SlitName, slitWidth);
	}
}

// calculate radiance
BOOL CMyOPTFile::CalculateRadiance(
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppSignal)
{
	return CalculateIrradiance(nValues, ppWaves, ppSignal, TRUE);

	/*
	long			nData;
	double		*	pWaves			= NULL;
	double		*	pSignal			= NULL;
	long			nCal;
	double		*	pWavesCal;
	double		*	pCalibration	= NULL;
	BOOL			fSuccess		= FALSE;
	long			i;
	double			calVal;
	double			FOV;
	double			fovFactor;
	double			PI		= atan(1.0) * 4.0;


	*nValues	= 0;
	*ppWaves	= NULL;
	*ppSignal	= NULL;

	if (!this->m_pInputInfo->getRadianceAvailable(&FOV))
	{
		return FALSE;
	}
	this->SetCalibrationMeasurement(FALSE);
	if (FOV == 0.0) FOV = 0.0017;
	fovFactor = 2.0 * PI * (1.0 - cos(FOV));
	if (this->GetNMScanData(&nData, &pWaves, &pSignal))
	{
		if (this->GetNMReferenceCalibration(pWaves[0], pWaves[nData-1], &nCal, &pWavesCal, &pCalibration))
		{
			if (nCal == nData)
			{
				*nValues	= nData;
				*ppWaves	= new double [nData];
				*ppSignal	= new double [nData];
				for (i=0; i<nData; i++)
				{
					(*ppWaves)[i]	= pWaves[i];
					calVal	= pCalibration[i] != 0.0 ? pCalibration[i] : 1.0;
					(*ppSignal)[i]	= pSignal[i] / calVal / fovFactor;
				}
				fSuccess	= TRUE;
			}
			delete [] pWavesCal;
			delete [] pCalibration;
		}
		delete [] pWaves;
		delete [] pSignal;
	}
	return fSuccess;
*/
}

// calculate Irradiance
BOOL CMyOPTFile::CalculateIrradiance(
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppSignal,
							BOOL			fRadianceCalc /*= FALSE*/)
{
	double		*	pWaves = NULL;
	double		*	pSignal = NULL;
	long			grating;
	TCHAR			szFilter[MAX_PATH];
	long			detector;
	long			iScan;
	long			tValues;
	double		*	tWaves;
	double		*	tSignals;
	long			nV;
	double		*	pX;
	double		*	pY;
	long			nCal;
	double		*	pCalWaves;
	double		*	pCalSignal;
	long			i;
	long			offset;
	double			calVal;
	double			bandPass;
	double			outputSlitWidth;
	CInterpolateValues	intValues;		// interpolate values utility
	double			calGainFactor;		// calibration gain factor
	double			minDataValue = this->GetMinSignalValue();
	double			DCOffset;			// DC signal offset
	// parameters used for Radiance calc
	double			fovFactor;
	double			FOV = 0.0;
	double			PI = atan(1.0) * 4.0;
	IDispatch	*	pdisp;
	BOOL			fRadiance = FALSE;

	*nValues = 0;
	*ppWaves = NULL;
	*ppSignal = NULL;
	if (fRadianceCalc)
	{
//		if (!this->m_pInputInfo->getRadianceAvailable(&FOV))
//		if (!this->GetRadianceAvailable())
//		{
//			return FALSE;
//		}
		if (this->GetInputInfo(&pdisp))
		{
			CInputInfo		inputInfo;
			inputInfo.doInit(pdisp);
			fRadiance = inputInfo.getRadianceAvailable(&FOV);

			pdisp->Release();
		}
		if (!fRadiance) return FALSE;

		if (FOV == 0.0) FOV = 0.0017;
//		fovFactor = 2.0 * PI * (1.0 - cos(FOV));
		// correction for the fov factor Sept 26, 2016 - above off by factor 4
		fovFactor = PI * (1.0 - cos(FOV)) / 2.0;
	}
	this->m_pSlitInfo->GetOutputSlitWidth(&outputSlitWidth);

	for (iScan = 0; iScan < this->m_nScans; iScan++)
	{
		if (this->m_ppScans[iScan]->GetNMInterpolatedData(&nV, &pX, &pY))
		{
			// determine bandpass
			bandPass = outputSlitWidth * this->m_ppScans[iScan]->Getdispersion();
			if (0.0 == bandPass) bandPass = 1.0;
			grating = this->m_ppScans[iScan]->Getgrating();
			this->m_ppScans[iScan]->Getfilter(szFilter, MAX_PATH);
			detector = this->m_ppScans[iScan]->GetDetector();
			if (this->m_pMyObject->FirerequestCalibrationScan(
				grating, szFilter,
				detector, &nCal, &pCalWaves, &pCalSignal))
			{
				tValues = (*nValues) + nV;
				tWaves = new double[tValues];
				tSignals = new double[tValues];
				if (NULL != (*ppWaves) && NULL != (*ppSignal))
				{
					for (i = 0; i < (*nValues); i++)
					{
						tWaves[i] = (*ppWaves)[i];
						tSignals[i] = (*ppSignal)[i];
					}
					offset = i;
					delete[](*ppWaves);
					delete[](*ppSignal);
				}
				else
				{
					offset = 0;
				}
				// apply the calibration
				intValues.SetXArray(nCal, pCalWaves);
				intValues.SetYArray(nCal, pCalSignal);
				// determine the calibration gain factor
				calGainFactor = this->DetermineCalibrationGainFactor(pX[nV / 2]);
				if (0.0 == calGainFactor) calGainFactor = 1.0;
				DCOffset = this->DetermineDCSignalOffset(pX[nV / 2]);
				for (i = 0; i < nV; i++)
				{
					tWaves[i + offset] = pX[i];
					pY[i] -= DCOffset;
					if (pY[i] < minDataValue) pY[i] = minDataValue;
					pY[i] /= bandPass;			// apply band pass
					calVal = intValues.interpolateValue(pX[i]);
					if (calVal < minDataValue) calVal = minDataValue;
					tSignals[i + offset] = pY[i] / calVal / calGainFactor;
					// apply FOV factor for  radiance calc
					if (fRadianceCalc)
						tSignals[i + offset] /= fovFactor;
				}
				// store
				(*nValues) = tValues;
				(*ppWaves) = tWaves;
				(*ppSignal) = tSignals;
			}
			delete[] pCalWaves;
			delete[] pCalSignal;
			delete[] pX;
			delete[] pY;
		}
	}
	return (*nValues) > 0;


/*
	long			nData;
	double		*	pWaves			= NULL;
	double		*	pSignal			= NULL;
	long			nCal;
	double		*	pWavesCal;
	double		*	pCalibration	= NULL;
	BOOL			fSuccess		= FALSE;
	long			i;
	double			calVal;

	*nValues	= 0;
	*ppWaves	= NULL;
	*ppSignal	= NULL;

	this->SetCalibrationMeasurement(FALSE);
	if (this->GetNMScanData(&nData, &pWaves, &pSignal))
	{
		if (this->GetNMReferenceCalibration(pWaves[0], pWaves[nData-1], &nCal, &pWavesCal, &pCalibration))
		{
			if (nCal == nData)
			{
				*nValues	= nData;
				*ppWaves	= new double [nData];
				*ppSignal	= new double [nData];
				for (i=0; i<nData; i++)
				{
					(*ppWaves)[i]	= pWaves[i];
					calVal	= pCalibration[i] != 0.0 ? pCalibration[i] : 1.0;
					(*ppSignal)[i]	= pSignal[i] / calVal;
				}
				fSuccess	= TRUE;
			}
			delete [] pWavesCal;
			delete [] pCalibration;
		}
		delete [] pWaves;
		delete [] pSignal;
	}
	return fSuccess;
*/
}

// get scan data interpolated to 1NM
BOOL CMyOPTFile::GetNMScanData(
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppSignal)
{

	double			*	pX	= NULL;
	long				nX	= 0;
	double			*	pY	= NULL;
	long				nY	= 0;
	CInterpolateValues	intValues;
	long				i;
	double				startWave;
	double				endWave;
	BOOL				fSuccess	= FALSE;
	
	*nValues	= 0;
	*ppWaves	= NULL;
	*ppSignal	= NULL;
	if (this->GetWavelengths(&nX, &pX))
	{
		if (this->GetSpectra(&nY, &pY, TRUE))
		{
			if (nX == nY)
			{
	//			this->SmoothArray(nX, pY);
				// interpolate to NM
				// find start and end wavelength
				startWave	= pX[0];
				endWave		= pX[nX-1];
//				*nValues	= (long) floor(endWave) - (long) floor(startWave) + 1;
				*nValues = (((long)floor(endWave + 0.5) - (long)floor(startWave + 0.5)) * 10) + 1;
				*ppWaves	= new double [*nValues];
				*ppSignal	= new double [*nValues];
				// interpolation
				intValues.SetXArray(nX, pX);
				intValues.SetYArray(nX, pY);
				for (i=0; i<(*nValues); i++)
				{
//					(*ppWaves)[i]	= floor(startWave) + (i * 1.0);
					(*ppWaves)[i]	= floor(startWave) + (i * 0.1);
					(*ppSignal)[i]	= intValues.interpolateValue((*ppWaves)[i]);
				}
				fSuccess	= TRUE;
			}
			delete [] pY;
		}
		delete [] pX;
	}
	return fSuccess;
}

BOOL CMyOPTFile::GetNMScanNormalizedData(			// scan data normalized by dispersion
	long		*	nValues,
	double		**	ppWaves,
	double		**	ppSignal)
{

	double			*	pX = NULL;
	long				nX = 0;
	double			*	pY = NULL;
	long				nY = 0;
	CInterpolateValues	intValues;
	long				i;
	double				startWave;
	double				endWave;
	BOOL				fSuccess = FALSE;

	*nValues = 0;
	*ppWaves = NULL;
	*ppSignal = NULL;
	if (this->GetWavelengths(&nX, &pX))
	{
		if (this->GetSpectra(&nY, &pY, TRUE))			// get the band pass corrected spectra
		{
			if (nX == nY)
			{
				// interpolate to NM
				// find start and end wavelength
				startWave = pX[0];
				endWave = pX[nX - 1];
				*nValues = (((long)floor(endWave + 0.5) - (long)floor(startWave + 0.5)) * 10) + 1;
				*ppWaves = new double[*nValues];
				*ppSignal = new double[*nValues];
				// interpolation
				intValues.SetXArray(nX, pX);
				intValues.SetYArray(nX, pY);
				for (i = 0; i<(*nValues); i++)
				{
					(*ppWaves)[i] = floor(startWave) + (i * 0.1);
					(*ppSignal)[i] = intValues.interpolateValue((*ppWaves)[i]);
				}
				fSuccess = TRUE;
			}
			delete[] pY;
		}
		delete[] pX;
	}
	return fSuccess;
}


// get Reference calibration data interpolated to 1 NM
BOOL CMyOPTFile::GetNMReferenceCalibration(
							double			startWave,
							double			endWave,
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppCal)
{
	double			*	pX	= NULL;
	long				nX	= 0;
	double			*	pY	= NULL;
	long				nY	= 0;
	CInterpolateValues	intValues;
	long				i;
	BOOL				fSuccess	= FALSE;
	double			*	tX			= NULL;
	double			*	tY			= NULL;
	long				tN			= 0;
	long				c;
	double				deltaX;
//	// lamp distance factor
	*nValues	= 0;
	*ppWaves	= NULL;
	*ppCal		= NULL;

	if (this->GetCalibrationData(&nX, &pX, &pY))
// request calibration data from client
//	if (this->m_pMyObject->FirerequestOpticalTransfer(&nX, &pX, &pY))
	{
		deltaX = floor(pX[1] + 0.5) - floor(pX[0] + 0.5);
		// obtain restricted array
		for (i = 0; i < nX; i++)
		{
			if (pX[i] >= (startWave- deltaX) && pX[i] <= (endWave + deltaX))
				tN++;
		}
		if (tN > 0)
		{
			fSuccess = TRUE;
			tX = new double[tN];
			tY = new double[tN];
			for (i = 0, c = 0; i < nX; i++)
			{
				if (pX[i] >= (startWave - deltaX) && pX[i] <= (endWave + deltaX))
				{
					tX[c] = pX[i];
					tY[c] = pY[i];		// *lampDistanceFactor;
					c++;
				}
			}
		}
		delete[] pX;
		delete[] pY;
	}
	if (fSuccess)
	{
		*nValues = (((long)floor(endWave + 0.5) - (long)floor(startWave + 0.5)) * 10) + 1;
		*ppWaves	= new double [*nValues];
		*ppCal		= new double [*nValues];
		// interpolation
		intValues.SetXArray(tN, tX);
		intValues.SetYArray(tN, tY);
		for (i=0; i<(*nValues); i++)
		{
			(*ppWaves)[i]	= floor(startWave) + (i * 0.01);
			(*ppCal)[i]		= intValues.interpolateValue((*ppWaves)[i]);
		}
		fSuccess	= TRUE;
		delete [] tX;
		delete [] tY;
	}
	return fSuccess;
}

// optical transfer function flag
BOOL CMyOPTFile::GetOpticalTransferFunction()
{
	return this->m_fOpticalTransferFunction;
}

// load optical transfer function file
BOOL CMyOPTFile::LoadOpticalTransferFile(
							LPCTSTR			szFileName)
{
	BOOL			fSuccess	= FALSE;
	// 
	if (NULL == szFileName) return FALSE;
	// check file extension
	if (0 == lstrcmpi(PathFindExtension(szFileName), TEXT("*.opt")))
	{
		fSuccess = this->LoadFromFile(szFileName);
	}
	else 
	{
		if (this->LoadFromFile(szFileName))
		{
			if (NULL != this->m_pCalibrationStandard)
			{
				fSuccess = this->FormOpticalTransfer();
			}
		}
	}
	if (fSuccess) this->SetOpticalTransferOn(TRUE);
	return fSuccess;
}

BOOL CMyOPTFile::FormOpticalTransfer()
{
	if (0 == this->m_nScans) return FALSE;
	long			iScan;
	long			nData;
	double		*	pWaves;
	double		*	pSignal;
	long			nCal;
	double		*	pWavesCal;
	double		*	pCalibration;
	BOOL			fSuccess		= FALSE;
	long			i;
	double			calVal;
	double			bandPass;
	double			outputSlitWidth;
	double			minSignalValue = this->GetMinSignalValue();		// minimum signal value
	double			DCSignalOffset;

//	this->DetermineAdditiveFactor();			// determine the additive factor required to make the signal data positive
//	// set am calibration flag
//	this->SetCalibrationMeasurement(TRUE);
/*
	if (this->GetNMScanNormalizedData(&nData, &pWaves, &pSignal))			// this signal corrected for bandpass
	{
//		// apply the reference calibration data - this signal is corrected for source distance
		if (this->GetNMReferenceCalibration(pWaves[0], pWaves[nData - 1], &nCal, &pWavesCal, &pCalibration))
		{
			if (nCal == nData)
			{
				fSuccess = TRUE;
				for (i = 0; i < nData; i++)
				{
					calVal = pCalibration[i] != 0.0 ? pCalibration[i] : 1.0;
					pSignal[i] /= calVal;
				}
			}
			delete[] pWavesCal;
			delete[] pCalibration;
		}
		// copy back to the grating scans
		if (fSuccess)
		{
			for (iScan = 0; iScan < this->m_nScans; iScan++)
			{
				this->m_ppScans[iScan]->GetWavelengthRange(&minX, &maxX);
				xRange[0] = minX - 0.01;
				if (iScan < (this->m_nScans - 1))
				{
					this->m_ppScans[iScan+1]->GetWavelengthRange(&minX, &maxX);
					xRange[1] = minX - 0.01;
				}
				else
				{
					xRange[1] = maxX + 10.0;
				}
				if (this->GetScanSection(nData, pWaves, pSignal, xRange[0], xRange[1], &nScan, &pWavesScan, &pSignalsScan))
				{
					this->m_ppScans[iScan]->SetDataArray(nScan, pWavesScan, pSignalsScan);
					delete[] pWavesScan;
					delete[] pSignalsScan;
				}
			}
		}
		// cleanup
		delete[] pWaves;
		delete[] pSignal;
*/
//	}
	// set am calibration flag
	this->m_pSlitInfo->GetOutputSlitWidth(&outputSlitWidth);
	this->SetCalibrationMeasurement(TRUE);
	for (iScan = 0; iScan < this->m_nScans; iScan++)
	{
		if (this->m_ppScans[iScan]->GetNMInterpolatedData(
			&nData, &pWaves, &pSignal))
		{
			// obtain the band pass
			bandPass = outputSlitWidth * this->m_ppScans[iScan]->Getdispersion();
			if (0.0 == bandPass) bandPass = 1.0;
			DCSignalOffset = this->DetermineDCSignalOffset(pWaves[nData / 2]);
			// apply the reference calibration
			if (this->GetNMReferenceCalibration(pWaves[0], pWaves[nData - 1], &nCal, &pWavesCal, &pCalibration))
			{
				if (nCal == nData)
				{
					fSuccess = TRUE;
					for (i = 0; i < nData; i++)
					{
						calVal = pCalibration[i] != 0.0 ? pCalibration[i] : 1.0;
						pSignal[i] -= DCSignalOffset;
						if (pSignal[i] < minSignalValue) pSignal[i] = minSignalValue;
						pSignal[i] /= bandPass;			// divide by band pass
						pSignal[i] /= calVal;			// divide by calibration
					}
					// write to the output
					this->m_ppScans[iScan]->SetDataArray(nData, pWaves, pSignal);
				}
				delete[] pWavesCal;
				delete[] pCalibration;
			}
			delete[] pWaves;
			delete[] pSignal;
		}
	}
	if (fSuccess) this->SetOpticalTransferOn(TRUE);
	return fSuccess;
}

// determine the additive factor for optical transfer calculation
void CMyOPTFile::DetermineAdditiveFactor()
{
	long		nArray;
	double	*	pSpectra;
	long		i;
	double		minValue = 0.0;
	this->m_addValue = 0.0;
	if (this->GetSpectra(&nArray, &pSpectra))
	{
		// find the minimum value
		minValue = pSpectra[0];
		for (i = 1; i < nArray; i++)
		{
			if (pSpectra[i] < minValue)
				minValue = pSpectra[i];
		}
		delete[] pSpectra;
	}
	if (minValue < 0.0)
	{
		this->m_addValue = 1.5 * (-minValue);
	}
}


BOOL CMyOPTFile::GetRefCalibration(
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppCal)
{
	return this->CalculateIrradiance(nValues, ppWaves, ppCal);
}

BOOL CMyOPTFile::SingleScanNMData(
							long			iScan,
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppSignal)
{
	double		bandPass;
	long		nX;
	double	*	pX;
	double	*	pY;
//	long		nY;
	long		i;
	double		outputSlitWidth;			//				= this->m_pSlitInfo->GetOutputSlitWidth();
	BOOL		fSuccess		= FALSE;
	CInterpolateValues	intValues;
	double		startWave;
	double		endWave;
	BOOL		fInterpolateOut = iScan < (this->m_nScans - 1);
	double		deltaX;

	*nValues	= NULL;
	*ppWaves	= NULL;
	*ppSignal	= NULL;
	if (iScan < 0 || iScan >= this->m_nScans) return FALSE;
	this->m_pSlitInfo->GetOutputSlitWidth(&outputSlitWidth);
	if (this->m_ppScans[iScan]->GetGratingScan(&nX, &pX, &pY))
	{
		fSuccess = TRUE;
		deltaX = floor(pX[1] + 0.5) - floor(pX[0] + 0.5);

		if (fInterpolateOut && deltaX >= 2.0)
		{
			long			tN = nX + 1;
			double		*	tX = new double[tN];
			double		*	tY = new double[tN];
			for (i = 0; i < nX; i++)
			{
				tX[i] = pX[i];
				tY[i] = pY[i];
			}
			tX[nX] = tX[nX - 1] + deltaX;
			tY[nX] = tY[nX - 1] + (tY[nX-1] - tY[nX-2]);
			delete[] pX;
			delete[] pY;
			pX = tX;
			pY = tY;
			nX = tN;
		}
		else
		{
			fInterpolateOut = FALSE;
		}
	}
	if (fSuccess)
	{
		bandPass = outputSlitWidth * this->m_ppScans[iScan]->Getdispersion();
		if (0.0 == bandPass) bandPass = 1.0;
		for (i=0; i<nX; i++)
		{
			pY[i] += this->m_addValue;			// additive value
			pY[i] /= bandPass;
		}
		deltaX = floor((pX[1] - pX[0]) + 0.5);
		// form output arrays
		startWave	= floor( pX[0] + 0.5);
		endWave		= floor(pX[nX-1] + 0.5);

		if (fInterpolateOut)
		{
			*nValues = (long)floor(endWave) - (long)floor(startWave);
		}
		else
		{
			*nValues = (long)floor(endWave) - (long)floor(startWave) + 1;
		}
		*ppWaves	= new double [*nValues];
		*ppSignal	= new double [*nValues];
		// interpolation
		intValues.SetXArray(nX, pX);
		intValues.SetYArray(nX, pY);
		for (i=0; i<(*nValues); i++)
		{
			(*ppWaves)[i]	= floor(startWave) + (i * 1.0);
			(*ppSignal)[i]	= intValues.interpolateValue((*ppWaves)[i]);
		}
		fSuccess = TRUE;
	}
	return fSuccess;
}

// get the time stamp
BOOL CMyOPTFile::GetTimeStamp(
							LPTSTR			szTimeStamp,
							UINT			nBufferSize)
{
	size_t			slen;
	BOOL			fSuccess	= FALSE;
	StringCchLength(this->m_szTimeStamp, MAX_PATH, &slen);
	if (slen > 0)
	{
		StringCchCopy(szTimeStamp, nBufferSize, this->m_szTimeStamp);
		fSuccess = TRUE;
	}
	else
	{
		szTimeStamp[0]	= '\0';
	}
	return fSuccess;
}

// set time stamp
void CMyOPTFile::SetTimeStamp()
{
	SYSTEMTIME			st;
	DATE				date;
	BSTR				bstrDate = NULL;
	LPTSTR				szTemp = NULL;

	// store the current time stamp
	GetSystemTime(&st);
	SystemTimeToVariantTime(&st, &date);
	VarBstrFromDate(date, LOCALE_USER_DEFAULT, 0, &bstrDate);
	if (NULL != bstrDate)
	{
		SHStrDup(bstrDate, &szTemp);
		StringCchCopy(this->m_szTimeStamp, MAX_PATH, szTemp);
		CoTaskMemFree((LPVOID)szTemp);
		SysFreeString(bstrDate);
	}
}


// initialize acquisition
BOOL CMyOPTFile::InitAcquisition(
							HWND			hwndParent,
							BOOL			fCalibration)
{	
//	HRESULT				hr;
	SYSTEMTIME			st;
	DATE				date;
	BSTR				bstrDate		= NULL;
	LPTSTR				szTemp			= NULL;
	BOOL				fSuccess		= FALSE;

	this->SetCalibrationMeasurement(fCalibration);
	// display setup
	if (this->ViewSetup(hwndParent))
	{
		fSuccess = TRUE;
		// store the current time stamp
		GetSystemTime(&st);
		SystemTimeToVariantTime(&st, &date);
		VarBstrFromDate(date, LOCALE_USER_DEFAULT, 0, &bstrDate);
		if (NULL != bstrDate)
		{
			SHStrDup(bstrDate, &szTemp);
			StringCchCopy(this->m_szTimeStamp, MAX_PATH, szTemp);
			CoTaskMemFree((LPVOID) szTemp);
			SysFreeString(bstrDate);
		}
	}
	// can commence acquisition
	return fSuccess;
}

// get the configuration file
BOOL CMyOPTFile::GetConfigFile(
							LPTSTR			szConfigFile,
							UINT			nBufferSize)
{
	size_t			slen;
	BOOL			fSuccess	= FALSE;
	StringCchLength(this->m_szConfigFile, MAX_PATH, &slen);
	if (slen > 0)
	{
		StringCchCopy(szConfigFile, nBufferSize, this->m_szConfigFile);
		fSuccess = TRUE;
	}
	return fSuccess;
}

// clear readonly
void CMyOPTFile::ClearReadonly()
{
	if (NULL != this->m_pDetectorInfo)
		this->m_pDetectorInfo->SetamReadOnly(FALSE);
//	if (NULL != this->m_pInputInfo)
//		this->m_pInputInfo->SetamReadOnly(FALSE);
	if (NULL != this->m_pSlitInfo)
		this->m_pSlitInfo->SetamReadOnly(FALSE);
	long			iScan;
	for (iScan=0; iScan < this->m_nScans; iScan++)
		this->m_ppScans[iScan]->SetamReadOnly(FALSE);
}

// find AD channel given wavelength
long CMyOPTFile::GetADChannel(
							double			wavelength)
{
	long		channel	= -1;
	if (NULL != this->m_pDetectorInfo)
	{
		channel = this->m_pDetectorInfo->findADChannel(wavelength);
	}
	return channel;
}

// file name
BOOL CMyOPTFile::GetFileName(
							LPTSTR			szFileName,
							UINT			nBufferSize)
{
	size_t			len;
	StringCchLength(this->m_szFilePath, MAX_PATH, &len);
	if (len > 0)
	{
		StringCchCopy(szFileName, nBufferSize, this->m_szFilePath);
		return TRUE;
	}
	else
	{
		szFileName[0]	= '\0';
		return FALSE;
	}
}

void CMyOPTFile::SetFileName(
							LPCTSTR			szFileName)
{
	StringCchCopy(this->m_szFilePath, MAX_PATH, szFileName);
}

BOOL CMyOPTFile::GetCalibrationData(
							long		*	nValues,
							double		**	ppWaves,
							double		**	ppSignals)
{
	long		nX;
	long		nY;
	double	*	pX;
	double	*	pY;
	long		i;
	BOOL		fSuccess = FALSE;
	double		lampDistanceFactor = this->GetLampDistanceFactor();
	*nValues	= 0;
	*ppWaves	= NULL;
	*ppSignals	= NULL;
	// if this is a calibration measurement then get the reference calibration
	if (this->GetCalibrationMeasurement())
	{
		if (NULL != this->m_pCalibrationStandard)
		{
			// make sure that the calibration standard is loaded
			if (!this->m_pCalibrationStandard->GetamLoaded())
			{
				this->m_pCalibrationStandard->SetamLoaded(TRUE);
			}
			if (this->m_pCalibrationStandard->Getwavelengths(&nX, &pX))
			{
				if (this->m_pCalibrationStandard->Getcalibration(&nY, &pY))
				{
					if (nX == nY)
					{
						fSuccess	= TRUE;
						*nValues	= nX;
						*ppWaves	= new double [nX];
						*ppSignals	= new double [nY];
						for (i=0; i<(*nValues); i++)
						{
							(*ppWaves)[i]	= pX[i];
							(*ppSignals)[i]	= pY[i] * lampDistanceFactor;
						}
					}
					delete [] pY;
				}
				delete [] pX;
			}
		}
	}
	else
	{
		// request calibration data from client
		fSuccess = this->m_pMyObject->FirerequestOpticalTransfer(nValues, ppWaves, ppSignals);
	}
	return fSuccess;
}

// single value
double CMyOPTFile::GetDataValue(
							double			wavelength)
{
	long		gratingScan;
	long		index;
	if (this->FindClosestValue(wavelength, &gratingScan, &index))
	{
		return this->m_ppScans[gratingScan]->getValue(index);
	}
	else
	{
		return 0.0;
	}
}

void CMyOPTFile::SetDataValue(
							double			wavelength,
							double			signal)
{
	long		gratingScan;
	long		index;
	if (this->FindClosestValue(wavelength, &gratingScan, &index))
	{
		this->m_ppScans[gratingScan]->setValue(index, signal);
	}
}

BOOL CMyOPTFile::FindClosestValue(
							double			wavelength,
							long		*	gratingScan,
							long		*	index)
{
	long			nScans		= this->GetNumGratingScans();
	long			iScan		= 0;
	long			arrSize;
	BOOL			fDone		= FALSE;
	long			ind;
	double			stepSize;

	*gratingScan	= -1;
	*index			= -1;
	while (iScan < nScans && !fDone)
	{
		arrSize = this->m_ppScans[iScan]->GetarraySize();
		stepSize = 1.0;
		ind = this->m_ppScans[iScan]->findIndex(wavelength, stepSize / 4.0);
		if (ind >=0)
		{
			fDone			= TRUE;
			*gratingScan	= iScan;
			*index			= ind;
		}
		if (!fDone) iScan++;
	}
	return fDone;
}

// form optical transfer file
BOOL CMyOPTFile::FormOpticalTransfer(
	LPCTSTR			szFilePath)
{
	// make sure that have reference calibration data
	this->SetCalibrationMeasurement(TRUE);
	if (!this->GetCalibrationMeasurement())
	{
		// try again - need to reload reference calibration
		this->Setup(GetApplicationWindow(), L"Calibration Standard");
		this->SetCalibrationMeasurement(TRUE);
		if (this->GetCalibrationMeasurement())
		{
			this->SaveToFile();		// resave this
		}
		else
		{
			return FALSE;
		}
	}
	// form optical transfer data
	if (!this->FormOpticalTransfer())
	{
		return FALSE;
	}
	// save to file
	return this->SaveToFile(szFilePath, TRUE);
}

// get scan data between given wavelengths
BOOL CMyOPTFile::GetScanSection(
	long		nData,
	double	*	pWaves,
	double	*	pSignals,
	double		minWave,
	double		maxWave,
	long	*	nOut,
	double	**	ppWavesOut,
	double	**	ppSignalsOut)
{
	long			i;
	long			c;
	double		*	tX;
	double		*	tY;
	long			nAlloc = nData;
	*nOut = 0;
	*ppWavesOut = NULL;
	*ppSignalsOut = NULL;
	tX = new double[nAlloc];
	tY = new double[nAlloc];
	for (i = 0, c = 0; i < nData; i++)
	{
		if (pWaves[i] >= minWave && pWaves[i] < maxWave)
		{
			tX[c] = pWaves[i];
			tY[c] = pSignals[i];
			c++;
		}
	}
	if (c > 0)
	{
		*nOut = c;
		*ppWavesOut = new double[*nOut];
		*ppSignalsOut = new double[*nOut];
		for (i = 0; i < (*nOut); i++)
		{
			(*ppWavesOut)[i] = tX[i];
			(*ppSignalsOut)[i] = tY[i];
		}
	}
	delete[] tX;
	delete[] tY;
	return (*nOut) > 0;
}

// smooth array
void CMyOPTFile::SmoothArray(
	long		nArray,
	double	*	pArray)
{
	long			i;
	double		*	tArray	= NULL;
	double			temp;
	tArray = new double[nArray];
	for (i = 0; i < nArray; i++)
		tArray[i] = 0.0;

	for (i = 0; i < nArray; i++)
	{
		if (i >= 2 && i < (nArray - 2))
		{
			temp = pArray[i - 2] + (4.0 * pArray[i - 1]) + (6.0 * pArray[i]) + (4.0 *pArray[i + 1]) + pArray[i + 2];
			tArray[i] = temp / 16.0;
		}
		else
		{
			if (i >= 1 && i < (nArray - 1))
			{
				tArray[i] = (pArray[i - 1] + (2.0 * pArray[i]) + pArray[i + 1]) / 4.0;
			}
			else
			{
				tArray[i] = pArray[i];
			}
		}
	}
	// copy back
	for (i = 0; i < nArray; i++)
	{
		pArray[i] = tArray[i];
	}
	delete[] tArray;
}

// lamp distance factor to apply to the reference calibration  - reference calibration at 50 cm
double CMyOPTFile::GetLampDistanceFactor()
{
	if (this->GetCalibrationMeasurement())
	{	
/*
		double			lampDistance = 1.0;
		IDispatch	*	pdisp;

		if (this->GetInputInfo(&pdisp))
		{
			CInputInfo		inputInfo;
			inputInfo.doInit(pdisp);
			lampDistance = inputInfo.GetlampDistance();
			pdisp->Release();
		}

//		double			lampDistance = this->m_pInputInfo->GetlampDistance();
		if (lampDistance < 1.0) lampDistance = 1.0;
		return (50.0 * 50.0) / (lampDistance * lampDistance);
		*/
		double expectedDistance = this->m_pCalibrationStandard->GetexpectedDistance();
		double sourceDistance = this->m_pCalibrationStandard->GetsourceDistance();
		if (sourceDistance < 1.0) sourceDistance = 1.0;
		return (expectedDistance * expectedDistance) / (sourceDistance * sourceDistance);
	}
	else
	{
		return 1.0;
	}

}

// comment string
BOOL CMyOPTFile::GetComment(
	LPTSTR			szComment,
	UINT			nBufferSize)
{
	size_t			slen;
	StringCchLength(this->m_szComment, MAX_PATH, &slen);
	if (slen > 0)
	{
		StringCchCopy(szComment, nBufferSize, this->m_szComment);
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

void CMyOPTFile::SetComment(
	LPCTSTR			szComment)
{
	StringCchCopy(this->m_szComment, MAX_PATH, szComment);
	Utils_OnPropChanged(this->m_pMyObject, DISPID_Comment);
}

// allow negative values flag
BOOL CMyOPTFile::GetAllowNegativeValues()
{
	return this->m_AllowNegativeValues;
}

void CMyOPTFile::SetAllowNegativeValues(
	BOOL			fAllowNegative)
{
	this->m_AllowNegativeValues = fAllowNegative;
	Utils_OnPropChanged(this->m_pMyObject, DISPID_AllowNegativeValues);
}

// extra scan value
void CMyOPTFile::ExtraScanValue(
	LPCTSTR			szTitle,
	double			wavelength,
	double			extraValue)
{
	if (this->m_nActiveScan >= 0 && this->m_nActiveScan < this->m_nScans)
	{
//		StringCchCopy(this->m_szExtraValue, MAX_PATH, szTitle);
		this->m_ppScans[this->m_nActiveScan]->setExtraValue(szTitle, wavelength, extraValue);
	}
/*
	TCHAR			szString[MAX_PATH];
	double		*	pValues = NULL;
	long			nValues = 0;
	if (this->GetExtraValues(szString, MAX_PATH, &nValues, &pValues))
	{
		delete[] pValues;
	}
*/
}

// number of extra values
long CMyOPTFile::GetNumberExtraValues(
	LPCTSTR		szTitle)
{
	long			nOut = 0;
	long			iScan;
	long			nScans = this->GetNumGratingScans();
	for (iScan = 0; iScan < nScans; iScan++)
	{
		nOut += (this->m_ppScans[iScan]->GetNumberExtraValues(szTitle));
	}
	return nOut;
}


BOOL CMyOPTFile::GetExtraValues(
	LPCTSTR			szExtraValue,
//	UINT			nBufferSize,
	long		*	nValues,
	double		**	ppExtraValues)
{
	*nValues = 0;
	*ppExtraValues = NULL;
//	double		arrValues[100000];			// make extra large buffer
	long		nArray;
	double	*	pArray;
	long		nOffset;
	long		i;
	long		nScans = this->GetNumGratingScans();
	long		iScan;
	long		nOut;						// number of output values
	long		nExtraValues = this->GetNumberExtraValues(szExtraValue);
	double	*	arrValues = NULL;
	BOOL		fSuccess = FALSE;
	if (0 == nExtraValues) return FALSE;
	arrValues = new double[nExtraValues];

//	StringCchCopy(szExtraValue, nBufferSize, this->m_szExtraValue);
	// loop over the scan data
	nOffset = 0;
	nOut = 0;
	for (iScan = 0; iScan < nScans; iScan++)
	{
		if (this->m_ppScans[iScan]->GetextraValues(szExtraValue, &nArray, &pArray))
		{
			for (i = 0; i < nArray; i++)
			{
				arrValues[nOffset + i] = pArray[i];
			}
			nOffset += nArray;
			nOut += nArray;
			delete[] pArray;
		}
	}
	if (nOut > 0)
	{
		*nValues = nOut;
		*ppExtraValues = new double[nOut];
		for (i = 0; i < nOut; i++)
			(*ppExtraValues)[i] = arrValues[i];
		fSuccess = TRUE;
	}
	delete[] arrValues;
	return fSuccess;
}

BOOL CMyOPTFile::GetExtraValueWaves(
	LPCTSTR			szExtraValue,
	long		*	nValues,
	double		**	ppWaves)
{
	*nValues = 0;
	*ppWaves = NULL;
//	double		arrValues[100000];			// make extra large buffer
	long		nArray;
	double	*	pArray;
	long		nOffset;
	long		i;
	long		nScans = this->GetNumGratingScans();
	long		iScan;
	long		nOut;						// number of output values
	long		nExtraValues = this->GetNumberExtraValues(szExtraValue);
	double	*	arrValues = NULL;
	BOOL		fSuccess = FALSE;
	if (0 == nExtraValues) return FALSE;
	arrValues = new double[nExtraValues];

											//	StringCchCopy(szExtraValue, nBufferSize, this->m_szExtraValue);
											// loop over the scan data
	nOffset = 0;
	nOut = 0;
	for (iScan = 0; iScan < nScans; iScan++)
	{
		if (this->m_ppScans[iScan]->GetextraValueWavelengths(szExtraValue, &nArray, &pArray))
		{
			for (i = 0; i < nArray; i++)
			{
				arrValues[nOffset + i] = pArray[i];
			}
			nOffset += nArray;
			nOut += nArray;
			delete[] pArray;
		}
	}
	if (nOut > 0)
	{
		*nValues = nOut;
		*ppWaves = new double[nOut];
		for (i = 0; i < nOut; i++)
			(*ppWaves)[i] = arrValues[i];
		fSuccess = TRUE;
	}
	delete[] arrValues;
	return fSuccess;
}

BOOL CMyOPTFile::HaveExtraValues()
{
/*
	TCHAR			szTitle[MAX_PATH];
	long			nValues;
	double		*	arrValues;
	if (this->GetExtraValues(szTitle, MAX_PATH, &nValues, &arrValues))
	{
		delete[] arrValues;
		return TRUE;
	}
	else
	{
		return FALSE;
	}
*/
	if (NULL != this->m_ppScans)
	{
		return this->m_ppScans[0]->GethaveExtraValue();
	}
	else
	{
		return FALSE;
	}
}

BOOL CMyOPTFile::GetExtraValueTitles(
	LPTSTR		**	ppszTitles,
	ULONG		*	nStrings)
{
	if (NULL != this->m_ppScans && this->m_nScans > 0)
	{
		return this->m_ppScans[0]->GetExtraValueTitles(ppszTitles, nStrings);
	}
	else
	{
		*ppszTitles = NULL;
		*nStrings = 0;
		return FALSE;
	}
}
long CMyOPTFile::GetNumberExtraScanArrays()
{
	ULONG			nTitles = 0;
	LPTSTR		*	pszTitles = NULL;
	ULONG			i;
	if (this->GetExtraValueTitles(&pszTitles, &nTitles))
	{
		// cleanup
		for (i = 0; i < nTitles; i++)
		{
			CoTaskMemFree((LPVOID)pszTitles[i]);
			pszTitles[i] = NULL;
		}
		CoTaskMemFree((LPVOID)pszTitles);
	}
	return nTitles;
}
BOOL CMyOPTFile::GetExtraValueTitle(
	long			index,
	LPTSTR			szTitle,
	UINT			nBufferSize)
{
	HRESULT			hr;
	ULONG			nTitles = 0;
	LPTSTR		*	pszTitles = NULL;
	ULONG			i;
	BOOL			fSuccess = FALSE;
	szTitle[0] = '\0';
	if (this->GetExtraValueTitles(&pszTitles, &nTitles))
	{
		hr = StringCchCopy(szTitle, nBufferSize, pszTitles[index]);
		fSuccess = SUCCEEDED(hr);
		// cleanup
		for (i = 0; i < nTitles; i++)
		{
			CoTaskMemFree((LPVOID)pszTitles[i]);
			pszTitles[i] = NULL;
		}
		CoTaskMemFree((LPVOID)pszTitles);
	}
	return fSuccess;
}

BOOL CMyOPTFile::GetExtraScanArray(
	long			index,
	long		*	nArray,
	double		**	ppwaves,
	double		**	ppvalues,
	LPTSTR			szTitle,
	UINT			nBufferSize)
{
	BOOL		fSuccess = FALSE;
	long		nX;
	long		nY;

	szTitle[0] = '\0';
	*nArray = 0;
	*ppwaves = NULL;
	*ppvalues = NULL;
	if (this->GetExtraValueTitle(index, szTitle, nBufferSize))
	{
		if (this->GetExtraValueWaves(szTitle, &nX, ppwaves))
		{
			if (this->GetExtraValues(szTitle, &nY, ppvalues))
			{
				if (nX == nY)
				{
					fSuccess = TRUE;
					*nArray = nX;
				}
				if (!fSuccess)
				{
					delete[] * ppvalues;
					*ppvalues = NULL;
				}
			}
			if (!fSuccess)
			{
				delete[] * ppwaves;
				*ppwaves = NULL;
			}
		}
	}
	return fSuccess;
}

// signal units
BOOL CMyOPTFile::GetSignalUnits(
	LPTSTR			szUnits,
	UINT			nBufferSize)
{
	StringCchCopy(szUnits, nBufferSize, this->m_szSignalUnits);
	return TRUE;
//	if (NULL != this->m_ppScans && this->m_nScans > 0)
//	{
//		return this->m_ppScans[0]->GetSignalUnits(szUnits, nBufferSize);
//	}
//	else
//	{
//		szUnits[0] = '\0';
//		return FALSE;
//	}
}
void CMyOPTFile::SetSignalUnits(
	LPCTSTR			szUnits)
{
	long		i;

	StringCchCopy(this->m_szSignalUnits, MAX_PATH, szUnits);


	for (i = 0; i < this->m_nScans; i++)
		this->m_ppScans[i]->SetSignalUnits(szUnits);
}

// default title
void CMyOPTFile::GetDefaultTitle(
	LPTSTR			szTitle,
	UINT			nBufferSize)
{
	StringCchCopy(szTitle, nBufferSize, this->m_szDefaultTitle);
}
void CMyOPTFile::SetDefaultTitle(
	LPCTSTR			szTitle)
{
	StringCchCopy(this->m_szDefaultTitle, MAX_PATH, szTitle);
}

// single scan data
BOOL CMyOPTFile::GetSingleGratingScan(
	long			grating,
	LPCTSTR			szFilter,
	long			detector,
	long		*	nArray,
	double		**	ppWaves,
	double		**	ppSignals)
{
	long		iScan = 0;
	BOOL		fDone = FALSE;
	BOOL		fSuccess = FALSE;
	long		nX;
	double	*	pX;
	long		nY;
	double	*	pY;

	*nArray = 0;
	*ppWaves = NULL;
	*ppSignals = NULL;

	while (iScan < this->m_nScans && !fDone)
	{
		fDone = this->m_ppScans[iScan]->CheckGratingAndFilter(
			grating, szFilter, detector);
		if (fDone)
		{
			if (this->m_ppScans[iScan]->Getwavelengths(&nX, &pX))
			{
				if (this->m_ppScans[iScan]->Getsignals(&nY, &pY))
				{
					if (nX == nY)
					{
						fSuccess = TRUE;
						*nArray = nX;
						*ppWaves = pX;
						*ppSignals = pY;
					}
					if (!fSuccess) delete[] pY;
				}
				if (!fSuccess) delete[] pX;
			}
		}
		if (!fDone) iScan++;
	}
	return fSuccess;
}

// export to CSV file
BOOL CMyOPTFile::ExportToCSV(
	LPCTSTR			szFilePath)
{
	HANDLE		hFile;
	BOOL		fSuccess = FALSE;
	long		iScan;					// scan index
	TCHAR		oneLine[MAX_PATH];		// one line to print
	size_t		slen;
	DWORD		dwNWritten;
	long		nX;
	double	*	pX;
	long		nY;
	double	*	pY;
	long		i;

	// create output file
	hFile = CreateFile(szFilePath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS,
		FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE != hFile)
	{
		fSuccess = TRUE;			// have created file
		// write the source path
		StringCchPrintf(oneLine, MAX_PATH, L"Source, %s\r\n", PathFindFileName(this->m_szFilePath));
		StringCbLength(oneLine, MAX_PATH * sizeof(TCHAR), &slen);
		WriteFile(hFile, (LPCVOID)oneLine, slen, &dwNWritten, NULL);
		// write the comment
		StringCchPrintf(oneLine, MAX_PATH, L"Comment, %s\r\n", this->m_szComment);
		StringCbLength(oneLine, MAX_PATH * sizeof(TCHAR), &slen);
		WriteFile(hFile, (LPCVOID)oneLine, slen, &dwNWritten, NULL);
		// loop over the data
		for (iScan = 0; iScan < this->m_nScans; iScan++)
		{
			if (this->m_ppScans[iScan]->Getwavelengths(&nX, &pX))
			{
				if (this->m_ppScans[iScan]->Getsignals(&nY, &pY))
				{
					if (nX == nY)
					{
						for (i = 0; i < nX; i++)
						{
							_stprintf_s(oneLine, L"%8.3f,%10.5g\r\n", pX[i], pY[i]);
							StringCbLength(oneLine, MAX_PATH * sizeof(TCHAR), &slen);
							WriteFile(hFile, (LPCVOID)oneLine, slen, &dwNWritten, NULL);
						}
					}
					delete[] pY;
				}
				delete[] pX;
			}
		}
		CloseHandle(hFile);
	}
	return fSuccess;
}

// determine detector and detector gain
BOOL CMyOPTFile::DetermineDetector(
	double			wavelength,
	LPTSTR			szDetector,
	LPTSTR			szDetectorGain)
{
//	long			index;
	long			iScan;
	BOOL			fSuccess = FALSE;
	long			detector;
	TCHAR			szModel[MAX_PATH];		// detector model
	BOOL			fDone = FALSE;
	long			nScans = this->GetNumGratingScans();
	double			minWave, maxWave;

	// determine the grating scan
//	if (this->FindClosestValue(wavelength, &iScan, &index))
	iScan = 0;
	while (iScan < nScans && !fDone)
	{
		if (this->m_ppScans[iScan]->GetWavelengthRange(&minWave, &maxWave))
		{
			fDone = wavelength > minWave && wavelength < maxWave;
		}
		if (!fDone) iScan++;
	}
	if (fDone)
	{
		detector = this->m_ppScans[iScan]->GetDetector();
		this->m_pDetectorInfo->GetdetectorModel(szModel, MAX_PATH);
		if (0 == lstrcmpi(szModel, L"UVSi/PbS"))
		{
			StringCchCopy(szDetector, MAX_PATH, 1 == detector ? L"UVSi" : L"PbS");
		}
		else if (0 == lstrcmpi(szModel, L"UVSi/PbSe"))
		{
			StringCchCopy(szDetector, MAX_PATH, 1 == detector ? L"UVSi" : L"PbSe");
		}
		else if (0 == lstrcmpi(szModel, L"PMT"))
		{
			StringCchCopy(szDetector, MAX_PATH, szModel);
		}
		this->m_pDetectorInfo->GetgainSetting(szDetectorGain, MAX_PATH);
	}
	return TRUE;
}

// determine the calibration gain factor
double CMyOPTFile::DetermineCalibrationGainFactor(
	double		wavelength)
{
	double		calGainFactor = this->m_pMyObject->FirerequestCalibrationGain(wavelength);
	double		dataGainFactor = this->DetermineDataGainFactor(wavelength);
	if (0.0 == dataGainFactor) dataGainFactor = 1.0;
	return calGainFactor / dataGainFactor;
}

// determine data gain factor
double CMyOPTFile::DetermineDataGainFactor(
	double		wavelength)
{
	DataGainFactor			dataGainFactor;
	this->DetermineDetectorInfo(wavelength, &dataGainFactor);
	return dataGainFactor.GetgainFactor();
}

// minimum signal value
double CMyOPTFile::GetMinSignalValue()
{
	DataGainFactor	dataGainFactor;
	return dataGainFactor.GetminRawDataValue();
}
// determine the DC signal offset
double CMyOPTFile::DetermineDCSignalOffset(
	double			wavelength)
{
	DataGainFactor			dataGainFactor;
	this->DetermineDetectorInfo(wavelength, &dataGainFactor);
	return dataGainFactor.GetDCOffset();
}

// determine detector info
BOOL CMyOPTFile::DetermineDetectorInfo(
	double		wavelength,
	DataGainFactor*	pDataGainFactor)
{
	double			gainFactor = 1.0;
	LPTSTR			szInfoString = NULL;		// AD info string
	LPTSTR			szRem;						// remainder string
	size_t			slen;
	size_t			i;
	TCHAR			szLine[MAX_PATH];
	BOOL			fDone;
	size_t			c;
	float			fval;
	short int		AcquireMode = -1;			// AcquireMode = 1 for MultiChannel, 2 for pulsed mode
	short int		Detector = -1;
	TCHAR			szDetector[MAX_PATH];
	TCHAR			szDetectorGain[MAX_PATH];

	if (NULL == pDataGainFactor) return FALSE;
	StringCchPrintf(szDetector, MAX_PATH, L"Unknown");
	StringCchPrintf(szDetectorGain, MAX_PATH, L"Unknown");
	if (this->GetADInfoString(&szInfoString))
	{
		// check if multi channel
		szRem = StrStrI(szInfoString, L"MultiChannel");
		if (NULL != szRem)
		{
			AcquireMode = 1;
			// check if using Lockin
			szRem = StrStrI(szInfoString, L"Lockin ON");
			if (NULL != szRem)
			{
				szRem = StrStrI(szInfoString, L"LockinGain");
				if (NULL != szRem)
				{
					StringCchLength(szRem, STRSAFE_MAX_CCH, &slen);
					if (slen > 0)
					{
						fDone = FALSE;
						i = 0;
						c = 0;
						while (i < slen && !fDone)
						{
							if ('\n' == szRem[i] || '\r' == szRem[i])
							{
								fDone = c > 0;
							}
							else
							{
								szLine[c] = szRem[i];
								c++;
							}
							if (!fDone) i++;
						}
						szLine[c] = '\0';
						if (1 == _stscanf_s(szLine, L"LockinGain %f", &fval))
						{
							gainFactor = fval;
						}
					}
				}
			}
		}
		else
		{
			AcquireMode = 2;
		}
		CoTaskMemFree((LPVOID)szInfoString);
	}
	//	InitVariantFromDouble(gainFactor, pVarResult);
	// determine gain and gain factor from wavelength
	this->DetermineDetector(wavelength, szDetector, szDetectorGain);
	if (0 == lstrcmpi(szDetector, L"PMT"))
	{
		Detector = 1;
	}
	else if (0 == lstrcmpi(szDetector, L"UVS") ||
		0 == lstrcmpi(szDetector, L"UVSi"))
	{
		Detector = 2;
	}
	else if (0 == lstrcmpi(szDetector, L"PbS"))
	{
		Detector = 3;
	}
	else if (0 == lstrcmpi(szDetector, L"PbSe"))
	{
		Detector = 4;
	}
	// set the values
	pDataGainFactor->Setdetector(Detector);
	pDataGainFactor->SetacquireMode(AcquireMode);
	pDataGainFactor->SetlockinGain((long)floor(gainFactor + 0.5));
	pDataGainFactor->SetdetectorGain(szDetectorGain);
	return TRUE;
}

// set optical transfer function flag on/off
void CMyOPTFile::SetOpticalTransferOn(
	BOOL		fSetOpticalTransfer)
{
	if (this->GetOpticalTransferFunction() == fSetOpticalTransfer) return;
	HRESULT				hr;
	IDispatch		*	pdisp;
	TCHAR				szUnits[MAX_PATH];
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		if (fSetOpticalTransfer)
		{
			this->m_fOpticalTransferFunction = fSetOpticalTransfer;
			StringCchCopy(szUnits, MAX_PATH, L"W/(cm^2)*nm");
		}
		else
		{
			this->m_fOpticalTransferFunction = fSetOpticalTransfer;
			StringCchCopy(szUnits, MAX_PATH, L"Volts");
		}
		Utils_SetStringProperty(pdisp, DISPID_SignalUnits, szUnits);
		pdisp->Release();
	}
}
