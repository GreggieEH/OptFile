#pragma once

using namespace std;

class CImpICalibrationStandard : public IDispatch
{
public:
	CImpICalibrationStandard();
	~CImpICalibrationStandard();
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
	HRESULT					GetamLoaded(
								VARIANT			*	pVarResult);
	HRESULT					SetamLoaded(
								DISPPARAMS		*	pDispParams);
	HRESULT					GetfileName(
								VARIANT			*	pVarResult);
	HRESULT					SetfileName(
								DISPPARAMS		*	pDispParams);
	HRESULT					Getwavelengths(
								VARIANT			*	pVarResult);
	HRESULT					Getcalibration(
								VARIANT			*	pVarResult);
	HRESULT					loadFromXML(
								DISPPARAMS		*	pDispParams,
								VARIANT			*	pVarResult);
	HRESULT					saveToXML(
								DISPPARAMS		*	pDispParams,
								VARIANT			*	pVarResult);
	HRESULT					loadExcelFile(
								DISPPARAMS		*	pDispParams,
								VARIANT			*	pVarResult);
	HRESULT					checkExcelFile(
								DISPPARAMS		*	pDispParams,
								VARIANT			*	pVarResult);
	BOOL					loadFromXML(
								IDispatch		*	xml);
	BOOL					saveToXML(
								IDispatch		*	xml);
	BOOL					GetfileName(
								LPTSTR			szFileName,
								UINT			nBufferSize);
	void					SetfileName(
								LPCTSTR			szfileName);
	BOOL					GetamLoaded();
	void					SetamLoaded(
								BOOL			amLoaded);
	BOOL					Getwavelengths(
								long		*	nValues,
								double		**	ppWaves);
	BOOL					Getcalibration(
								long		*	nValues,
								double		**	ppcalibration);
	BOOL					loadExcelFile(
								LPCTSTR			szFileName);
	BOOL					checkExcelFile(
								LPCTSTR			szFileName);
	void					ClearData();
	HRESULT					GetexpectedDistance(
								VARIANT		*	pVarResult);
	HRESULT					SetexpectedDistance(
								DISPPARAMS	*	pDispParams);
	HRESULT					GetsourceDistance(
								VARIANT		*	pVarResult);
	HRESULT					SetsourceDistance(
								DISPPARAMS	*	pDispParams);
	void					AddValue(
								double			x,
								double			y);
private:
	class CDataPoint
	{
	public:
		CDataPoint();
		CDataPoint(const CDataPoint& src);
		CDataPoint(double x, double y);
		~CDataPoint();
		const CDataPoint&	operator=(const CDataPoint& src);
		inline double		GetX();
		inline void			SetX(
								double		x);
		inline double		GetY();
		inline void			SetY(
								double		y);
	private:
		double				m_X;
		double				m_Y;
	};
	ULONG				m_cRefs;
	TCHAR				m_szFileName[MAX_PATH];
	vector<CDataPoint>	m_Data;
//	long				m_nArray;
//	double			*	m_pWaves;
//	double			*	m_pCalibration;
	double				m_expectedDistance;
	double				m_sourceDistance;
};

#include "ImpICalibrationStandard.inl"

