#pragma once
#include "dispobject.h"

class CGratingScanInfo : public CDispObject
{
public:
	CGratingScanInfo(void);
	virtual ~CGratingScanInfo(void);
	BOOL		InitNew();
	BOOL		loadFromXML(
					IDispatch	*	xml);
	BOOL		saveToXML(
					IDispatch	*	xml);
	BOOL		GetamReadOnly();
	void		SetamReadOnly(
					BOOL			amReadOnly);
	long		Getgrating();
	void		Setgrating(
					long			grating);
	BOOL		Getfilter(
					LPTSTR			szFilter,
					UINT			nBufferSize);
	void		Setfilter(
					LPCTSTR			szFilter);
	long		GetDetector();
	void		SetDetector(
					long			detector);
	double		Getdispersion();
	void		Setdispersion(
					double			dispersion);
	BOOL		Getwavelengths(
					long		*	nValues,
					double		**	ppWaves);
	BOOL		Getsignals(
					long		*	nValues,
					double		**	ppSignals);
	void		addData(
					double			wavelength,
					double			signal);
	BOOL		GetnodeName(
					LPTSTR			sznodeName,
					UINT			nBufferSize);
	void		SetnodeName(
					LPCTSTR			sznodeName);
	BOOL		GetamDirty();
	void		clearDirty();
	BOOL		GetGratingScan(
					long		*	nValues,
					double		**	ppWaves,
					double		**	ppSignals);
	long		GetarraySize();
	BOOL		CheckGratingAndFilter(
					long			grating,
					LPCTSTR			szFilter,
					long			detector);
	BOOL		SetDataArray(
					long			nValues,
					double		*	pWaves,
					double		*	pSignals);
//	long		findClosestIndex(
//					double			wavelength);
	double		getValue(
					long			index);
	void		setValue(
					long			index,
					double			signal);
	BOOL		GetWavelengthRange(
					double		*	minWave,
					double		*	maxWave);
	long		findIndex(
					double			wavelength,
					double			tolerance);
	BOOL		GethaveExtraValue();
	void		SethaveExtraValue(
					BOOL			fhaveExtraValue);
	void		setExtraValue(
					LPCTSTR			szTitle,
					double			wavelength,
					double			extraValue);
	BOOL		GetextraValues(
					LPCTSTR			szTitle,
					long		*	nValues,
					double		**	ppextraValues);
	BOOL		GetextraValueWavelengths(
					LPCTSTR			szTitle,
					long		*	nValues,
					double		**	ppWaves);
	BOOL		GetExtraValueTitles(
					LPTSTR		**	ppszTitles,
					ULONG		*	nStrings);
	long		GetNumberExtraValues(
					LPCTSTR			szTitle);
	// signal units
	BOOL		GetSignalUnits(
					LPTSTR			szUnits,
					UINT			nBufferSize);
	void		SetSignalUnits(
					LPCTSTR			szUnits);
	// interpolated data to 0.1 nm
	BOOL		GetNMInterpolatedData(
					long		*	nValues,
					double		**	ppWaves,
					double		**	ppSignals);
protected:
	void		SetLogging();
	void		OnLogString();
	void		ClearLogging();
	void		SetLoadSpectralData();
	void		ClearLoadSpectralData();
	void		OnHaveSpectralData();
	BOOL		LoadSpectralData(
					LPTSTR			szString);
	void		SetWavelengths(
					long			nArray,
					double		*	arr);
	void		SetSignals(
					long			nArray,
					double		*	arr);
private:
	class CImpIDispatch : public IDispatch
	{
	public:
		CImpIDispatch(CGratingScanInfo* pGratingScanInfo, LPCTSTR szName);
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
	private:
		CGratingScanInfo*	m_pGratingScanInfo;
		ULONG				m_cRefs;
		TCHAR				m_szName[MAX_PATH];
	};
	friend CImpIDispatch;
};
