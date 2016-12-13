#pragma once

class CMyOPTFile;

class CMyObject : public IUnknown
{
public:
	CMyObject(IUnknown * pUnkOuter);
	~CMyObject(void);
	// IUnknown methods
	STDMETHODIMP				QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)		AddRef();
	STDMETHODIMP_(ULONG)		Release();
	// initialization
	HRESULT						Init();
	// sink events
	BOOL						FireRequestSlitWidth(
									LPCTSTR			szSlitName,
									double		*	slitWidth);
	BOOL						FireRequestGetInputPort(
									LPTSTR			szInputPort,
									UINT			nBufferSize);
	BOOL						FireRequestSetInputPort(
									LPCTSTR			szInputPort);
	BOOL						FireRequestGetOutputPort(
									LPTSTR			szOutputPort,
									UINT			nBufferSize);
	BOOL						FireRequestSetOutputPort(
									LPCTSTR			szOutputPort);
	// __clsIDataSet sink
	double						FirerequestDispersion(
									long			grating);
	BOOL						FirerequestOpticalTransfer(
									long		*	nValues,
									double		**	ppWaves,
									double		**	ppSignals);
	BOOL						FirerequestCalibrationScan(
									short int		Grating,
									LPCTSTR			filter,
									short int		detector,
									long		*	nValues,
									double		**	ppWaves,
									double		**	ppSignals);
	double						FirerequestCalibrationGain(
									double			wavelength);
protected:
	HRESULT						GetClassInfo(
									ITypeInfo	**	ppTI);
	HRESULT						GetRefTypeInfo(
									LPCTSTR			szInterface,
									ITypeInfo	**	ppTypeInfo);
	HRESULT						Init__clsIDataSet();
private:
	class CImpIDispatch : public IDispatch
	{
	public:
		CImpIDispatch(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIDispatch();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount( 
									PUINT			pctinfo);
		STDMETHODIMP			GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo);
		STDMETHODIMP			GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId);
		STDMETHODIMP			Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr); 
	protected:
		HRESULT					GetDetectorInfo(
									VARIANT		*	pVarResult);
		HRESULT					GetInputInfo(
									VARIANT		*	pVarResult);
		HRESULT					GetSlitInfo(
									VARIANT		*	pVarResult);
		HRESULT					GetNumGratingScans(
									VARIANT		*	pVarResult);
		HRESULT					GetRadianceAvailable(
									VARIANT		*	pVarResult);
		HRESULT					GetCalibrationStandard(
									VARIANT		*	pVarResult);
		HRESULT					SetCalibrationStandard(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetCalibrationMeasurement(
									VARIANT		*	pVarResult);
		HRESULT					SetCalibrationMeasurement(
									DISPPARAMS	*	pDispParams);
		HRESULT					Clone(
									VARIANT		*	pVarResult);
		HRESULT					Setup(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetGratingScan(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					InitNew(
									VARIANT		*	pVarResult);
		HRESULT					ReadConfig(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					ReadINI(
									DISPPARAMS	*	 pDispParams);
		HRESULT					WriteINI(
									DISPPARAMS	*	pDispParams);
		HRESULT					SetExp(
									DISPPARAMS	*	pDispParams);
		HRESULT					AddValue(
									DISPPARAMS	*	pDispParams);
		HRESULT					SaveToFile(
									DISPPARAMS	*	pDispParams);
		HRESULT					LoadFromFile(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					ClearScanData();
		HRESULT					SetSlitInfo(
									DISPPARAMS	*	pDispParams);
		HRESULT					SelectCalibrationStandard(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					ClearCalibrationStandard();
		HRESULT					GetOpticalTransferFunction(
									VARIANT		*	pVarResult);
		HRESULT					LoadOpticalTransferFile(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					DisplayRadiance(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetTimeStamp(
									VARIANT		*	pVarResult);
		HRESULT					InitAcquisition(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetConfigFile(
									VARIANT		*	pVarResult);
		HRESULT					ClearReadonly();
		HRESULT					GetADChannel(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SaveToString(
									VARIANT		*	pVarResult);
		HRESULT					LoadFromString(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					FormOpticalTransferFile(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetADInfoString(
									VARIANT		*	pVarResult);
		HRESULT					SetADInfoString(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetComment(
									VARIANT		*	pVarResult);
		HRESULT					SetComment(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetAllowNegativeValues(
									VARIANT		*	pVarResult);
		HRESULT					SetAllowNegativeValues(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetExtraValues(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SetTimeStamp();
		HRESULT					GetExtraValueTitles(
									VARIANT		*	pVarResult);
		HRESULT					GetSignalUnits(
									VARIANT		*	pVarResult);
		HRESULT					SetSignalUnits(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetADGainFactor(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					RecallSettings(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SaveSettings(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetNMReferenceCalibration(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
		ITypeInfo			*	m_pTypeInfo;
	};
	class CImpIProvideClassInfo2 : public IProvideClassInfo2
	{
	public:
		CImpIProvideClassInfo2(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIProvideClassInfo2();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IProvideClassInfo method
		STDMETHODIMP			GetClassInfo(
									ITypeInfo	**	ppTI);  
		// IProvideClassInfo2 method
		STDMETHODIMP			GetGUID(
									DWORD			dwGuidKind,  //Desired GUID
									GUID		*	pGUID);       
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpIConnectionPointContainer : public IConnectionPointContainer
	{
	public:
		CImpIConnectionPointContainer(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIConnectionPointContainer();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IConnectionPointContainer methods
		STDMETHODIMP			EnumConnectionPoints(
									IEnumConnectionPoints **ppEnum);
		STDMETHODIMP			FindConnectionPoint(
									REFIID			riid,  //Requested connection point's interface identifier
									IConnectionPoint **ppCP);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImp_clsIDataSet : public IDispatch
	{
	public:
		CImp_clsIDataSet(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImp_clsIDataSet();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount( 
									PUINT			pctinfo);
		STDMETHODIMP			GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo);
		STDMETHODIMP			GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId);
		STDMETHODIMP			Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr); 
		HRESULT					MyInit();
	protected:
		HRESULT					ReadDataFile(
									VARIANT		*	pVarResult);
		HRESULT					WriteDataFile(
									VARIANT		*	pVarResult);
		HRESULT					GetWavelengths(
									VARIANT		*	pVarResult);
		HRESULT					GetSpectra(
									VARIANT		*	pVarResult);
		HRESULT					AddValue(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					ViewSetup(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetArraySize(
									VARIANT		*	pVarResult);
		HRESULT					SetCurrentExp(
									DISPPARAMS	*	pDispParams);
		HRESULT					verifyCalibration(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetRadianceAvailable(
									VARIANT		*	pVarResult);
		HRESULT					GetIrradianceAvailable(
									VARIANT		*	pVarResult);
		HRESULT					calcRadiance(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetAnalysisMode(
									VARIANT		*	pVarResult);
		HRESULT					SetAnalysisMode(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetCreateTime(
									VARIANT		*	pVarResult);
		HRESULT					GetfileName(
									VARIANT		*	pVarResult);
		HRESULT					SetfileName(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetamCalibration(
									VARIANT		*	pVarResult);
		HRESULT					SetamCalibration(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetDataValue(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SetDataValue(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetComment(
									VARIANT		*	pVarResult);
		HRESULT					SetComment(
									DISPPARAMS	*	pDispParams);
		HRESULT					ExtraScanData(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetDefaultUnits(
									VARIANT		*	pVarResult);
		HRESULT					SetDefaultUnits(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetDefaultTitle(
									VARIANT		*	pVarResult);
		HRESULT					SetDefaultTitle(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetCountsAvailable(
									VARIANT		*	pVarResult);
		HRESULT					SetCountsAvailable(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetNumberExtraScanArrays(
									VARIANT		*	pVarResult);
		HRESULT					GetExtraScanArray(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetGratingScan(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					ExportToCSV(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetADGainFactor(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
		ITypeInfo			*	m_pTypeInfo;
		short int				m_AnalysisMode;
		DISPID					m_dispidReadDataFile;
		DISPID					m_dispidWriteDataFile;
		DISPID					m_dispidGetWavelengths;
		DISPID					m_dispidGetSpectra;
		DISPID					m_dispidAddValue;
		DISPID					m_dispidViewSetup;
		DISPID					m_dispidGetArraySize;
		DISPID					m_dispidSetCurrentExp;
		DISPID					m_dispidverifyCalibration;
		DISPID					m_dispidRadianceAvailable;
		DISPID					m_dispidIrradianceAvailable;
		DISPID					m_dispidcalcRadiance;
		DISPID					m_dispidAnalysisMode;
		DISPID					m_dispidGetCreateTime;
		DISPID					m_dispidfileName;
		DISPID					m_dispidamCalibration;
		DISPID					m_dispidGetDataValue;
		DISPID					m_dispidSetDataValue;
		DISPID					m_dispidComment;
		DISPID					m_dispidExtraScanData;
		DISPID					m_dispidDefaultUnits;
		DISPID					m_dispidDefaultTitle;
		DISPID					m_dispidCountsAvailable;
		DISPID					m_dispidNumberExtraScanArrays;
		DISPID					m_dispidGetExtraScanArray;
		DISPID					m_dispidGetGratingScan;
		DISPID					m_dispidExportToCSV;
		DISPID					m_dispidGetADGainFactor;
	};
	class CImpIPersistStreamInit : public IPersistStreamInit
	{
	public:
		CImpIPersistStreamInit(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIPersistStreamInit();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IPersist method
		STDMETHODIMP			GetClassID(
									CLSID		*	pClassID);
		// IPersistStreamInit methods
		STDMETHODIMP			GetSizeMax(
									ULARGE_INTEGER *pCbSize);
		STDMETHODIMP			InitNew();
		STDMETHODIMP			IsDirty();
		STDMETHODIMP			Load(
									LPSTREAM		pStm);
		STDMETHODIMP			Save(
									LPSTREAM		pStm,
									BOOL			fClearDirty);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpIPersistFile : public IPersistFile
	{
	public:
		CImpIPersistFile(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIPersistFile();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IPersist method
		STDMETHODIMP			GetClassID(
									CLSID		*	pClassID);
		// IPersistFile methods
		STDMETHODIMP			GetCurFile(
									LPOLESTR	*	ppszFileName);
		STDMETHODIMP			IsDirty();
		STDMETHODIMP			Load(
									LPCOLESTR		pszFileName,
									DWORD			dwMode);
		STDMETHODIMP			Save(
									LPCOLESTR		pszFileName,
									BOOL			fRemember);
		STDMETHODIMP			SaveCompleted(
									LPCOLESTR		pszFileName);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
		LPTSTR					m_szFileName;
		BOOL					m_fNoScribble;
	};
	class CImpISummaryInfo : public IDispatch
	{
	public:
		CImpISummaryInfo(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpISummaryInfo();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount(
									PUINT			pctinfo);
		STDMETHODIMP			GetTypeInfo(
									UINT			iTInfo,
									LCID			lcid,
									ITypeInfo	**	ppTInfo);
		STDMETHODIMP			GetIDsOfNames(
									REFIID			riid,
									OLECHAR		**  rgszNames,
									UINT			cNames,
									LCID			lcid,
									DISPID		*	rgDispId);
		STDMETHODIMP			Invoke(
									DISPID			dispIdMember,
									REFIID			riid,
									LCID			lcid,
									WORD			wFlags,
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult,
									EXCEPINFO	*	pExcepInfo,
									PUINT			puArgErr);
		HRESULT					MyInit();
	protected:
		HRESULT					GetComment(
									VARIANT		*	pVarResult);
		HRESULT					GetXData(
									VARIANT		*	pVarResult);
		HRESULT					GetYData(
									VARIANT		*	pVarResult);
		HRESULT					GetAcquisitionDate(
									VARIANT		*	pVarResult);
		HRESULT					GetXRange(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetYRange(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
		ITypeInfo			*	m_pTypeInfo;
		DISPID					m_dispidComment;
		DISPID					m_dispidXData;
		DISPID					m_dispidYData;
		DISPID					m_dispidAcquisitionDate;
		DISPID					m_dispidXRange;
		DISPID					m_dispidYRange;
	};
	friend CImpIDispatch;
	friend CImpIConnectionPointContainer;
	friend CImpIProvideClassInfo2;
	friend CImp_clsIDataSet;
	friend CImpIPersistStreamInit;
	friend CImpIPersistFile;
	friend CImpISummaryInfo;
	// data members
	CImpIDispatch			*	m_pImpIDispatch;
	CImpIConnectionPointContainer*	m_pImpIConnectionPointContainer;
	CImpIProvideClassInfo2	*	m_pImpIProvideClassInfo2;
	CImp_clsIDataSet		*	m_pImp_clsIDataSet;
	CImpIPersistStreamInit	*	m_pImpIPersistStreamInit;
	CImpIPersistFile		*	m_pImpIPersistFile;
	CImpISummaryInfo		*	m_pImpISummaryInfo;
	// outer unknown for aggregation
	IUnknown				*	m_pUnkOuter;
	// object reference count
	ULONG						m_cRefs;
	// connection points array
	IConnectionPoint		*	m_paConnPts[MAX_CONN_PTS];
	// _clsIDataSet interface ID
	IID							m_iid_clsIDataSet;
	// opt file
	CMyOPTFile				*	m_pMyOPTFile;
	// __clsIDataSet event interface
	IID							m_iid__clsIDataSet;
	DISPID						m_dispidrequestDispersion;
	DISPID						m_dispidrequestOpticalTransfer;
	DISPID						m_dispidrequestCalibrationScan;
	DISPID						m_dispidrequestCalibrationGain;
	// summary info type info
	IID							m_iidISummayInfo;
};
