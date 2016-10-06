#pragma once

class CSciXML;
class CParseLine;

using namespace std;

class CImpGratingScanInfo : public IDispatch
{
public:
	CImpGratingScanInfo();
	~CImpGratingScanInfo();
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
	HRESULT					loadFromXML(
								DISPPARAMS	*	pDispParams,
								VARIANT		*	pVarResult);
	HRESULT					saveToXML(
								DISPPARAMS	*	pDispParams,
								VARIANT		*	pVarResult);
	HRESULT					addData(
								DISPPARAMS	*	pDispParams);
	HRESULT					ClearData();
	HRESULT					clearDirty();
	HRESULT					findIndex(
								DISPPARAMS	*	pDispParams,
								VARIANT		*	pVarResult);
	HRESULT					GetNodeName(
								VARIANT		*	pVarResult);
	HRESULT					SetNodeName(
								DISPPARAMS	*	pDispParams);
	HRESULT					GetGrating(
								VARIANT		*	pVarResult);
	HRESULT					SetGrating(
								DISPPARAMS	*	pDispParams);
	HRESULT					GetFilter(
								VARIANT		*	pVarResult);
	HRESULT					SetFilter(
								DISPPARAMS	*	pDispParams);
	HRESULT					GetDispersion(
								VARIANT		*	pVarResult);
	HRESULT					SetDispersion(
								DISPPARAMS	*	pDispParams);
	HRESULT					GetWavelengths(
								VARIANT		*	pVarResult);
	HRESULT					SetWavelengths(
								DISPPARAMS	*	pDispParams);
	HRESULT					GetSignals(
								VARIANT		*	pVarResult);
	HRESULT					SetSignals(
								DISPPARAMS	*	pDispParams);
	HRESULT					GetAmDirty(
								VARIANT		*	pVarResult);
	BOOL					loadFromXML(
								CSciXML		*	pSciXML);
	BOOL					LoadSpectralData(
								LPTSTR			szString);
	BOOL					FormLine(
								LPTSTR		*	szString,
								LPTSTR			szOneLine,
								UINT			nBufferSize);
	BOOL					ReadTitles(
								LPCTSTR			szOneLine);
	BOOL					ReadOneLine(
								LPTSTR			szOneLine,
								double		*	x,
								double		*	y);
	BOOL					saveToXML(
								CSciXML		*	pSciXML);
	BOOL					writeSpectralData(
								LPTSTR		*	szSpectralData);
	BOOL					WriteTitle(
								LPTSTR			szTitle,
								UINT			nBufferSize);
	void					WriteData(
								long			index,
								LPTSTR			szOneLine,
								UINT			nBufferSize);
	HRESULT					GetarraySize(
								VARIANT		*	pVarResult);
	HRESULT					getValue(
								DISPPARAMS	*	pDispParams,
								VARIANT		*	pVarResult);
	HRESULT					setValue(
								DISPPARAMS	*	pDispParams);
	HRESULT					setExtraValue(
								DISPPARAMS	*	pDispParams);
//	HRESULT					GetextraValues(
//								VARIANT		*	pVarResult);
//	void					LoadExtraValues(
//								LPTSTR			szExtraValues);
	void					AddExtraValue(
								LPCTSTR			szTitle,
								double			wavelength,
								double			extraValue);
//	BOOL					WriteExtraValues(
//								LPTSTR		*	szExtraValues);
	HRESULT					GethaveExtraValue(
								VARIANT		*	pVarResult);
	void					ClearExtraValues();
	HRESULT					GetExtraValuesTitles(
								VARIANT			*	pVarResult);
	HRESULT					GetSignalUnits(
								VARIANT			*	pVarResult);
	HRESULT					SetSignalUnits(
								DISPPARAMS		*	pDispParams);
	HRESULT					GetExtraValueComponent(
								DISPPARAMS		*	pDispParams,
								VARIANT			*	pVarResult);
	HRESULT					GetExtraValueWavelengths(
								DISPPARAMS		*	pDispParams,
								VARIANT			*	pVarResult);
	HRESULT					GetDetector(
								VARIANT			*	pVarResult);
	HRESULT					SetDetector(
								DISPPARAMS		*	pDispParams);
private:
	class CSpectralData
	{
	public:
		CSpectralData();
		CSpectralData(const CSpectralData& src);
		CSpectralData(double wavelength, double signal);			
		~CSpectralData();
		const CSpectralData&	operator=(const CSpectralData& src);
		double				GetWavelength();
		void				SetWavelength(
								double			wavelength);
		double				GetSignal();
		void				AddSignal(
								double			signal);
	private:
		double				m_Wavelength;
		vector<double>		m_signals;
	};
	class CExtraValues
	{
	public:
		CExtraValues();
		~CExtraValues();
		void				SetTitle(
								LPCTSTR				szTitle);
		BOOL				GetTitle(
								LPTSTR				szTitle,
								UINT				nBufferSize);
		void				AddValue(
								double				wavelength,
								double				newValue);
		BOOL				doFindValue(
								double				wavelength,
								double				tolerance,
								double			*	value);
		long				GetNValues();
		BOOL				GetWavelengths(
								long			*	nValues,
								double			**	ppWavelengths);
		BOOL				GetExtraValues(
								long			*	nValues,
								double			**	ppExtraValues);
	protected:
		long				FindIndex(
								double				wavelength,
								double				tolerance);
	private:
		TCHAR				m_szTitle[MAX_PATH];
		vector<CSpectralData>	m_ExtraValues;
	};
	friend CExtraValues;
	BOOL					FindExtraValueComponent(
								LPCTSTR				szTitle,
								CExtraValues	**	ppExtraValues);

	ULONG					m_cRefs;					// object reference count
	TCHAR					m_szNodeName[MAX_PATH];		// object node name
	long					m_grating;
	TCHAR					m_szFilter[MAX_PATH];
	double					m_dispersion;
	long					m_detector;
	// spectral data
	long					m_nWaves;
	double				*	m_pWavelengths;
	long					m_nSignals;
	double				*	m_pSignals;
	TCHAR					m_szSignal[MAX_PATH];		// signal title
	// extra values
	long					m_nExtraValues;
	CExtraValues		**	m_ppExtraValues;
	// dirty flag
	BOOL					m_amDirty;

};

