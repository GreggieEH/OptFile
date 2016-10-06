#pragma once

class CDispObject
{
public:
	CDispObject(void);
	virtual ~CDispObject(void);
	BOOL			doInit(
						LPCTSTR			szProgID);
	BOOL			doInit(
						IDispatch	*	pdisp);
	BOOL			getMyObject(
						IDispatch	**	ppdisp);
	// properties
	HRESULT			getProperty(
						LPCTSTR			szProp,
						VARIANT		*	Value);
	HRESULT			setProperty(
						LPCTSTR			szProp,
						VARIANTARG	*	Value);
	BOOL			getBoolProperty(
						LPCTSTR			szProp);
	void			setBoolProperty(
						LPCTSTR			szProp,
						BOOL			bValue);
	long			getLongProperty(
						LPCTSTR			szProp);
	void			setLongProperty(
						LPCTSTR			szProp,
						long			lval);
	double			getDoubleProperty(
						LPCTSTR			szProp);
	void			setDoubleProperty(
						LPCTSTR			szProp,
						double			dval);
	BOOL			getStringProperty(
						LPCTSTR			szProp,
						LPTSTR			szString,
						UINT			nBufferSize);
	BOOL			getStringProperty(
						LPCTSTR			szProp,
						LPTSTR		*	szString);
	void			setStringProperty(
						LPCTSTR			szProp,
						LPCTSTR			szString);
	HRESULT			InvokeMethod(
						LPCTSTR			szMethod,
						VARIANTARG	*	pvarg, 
						int				cArgs,
						VARIANT		*	pVarResult);
	HRESULT			InvokeMethod(
						DISPID			dispid,
						VARIANTARG	*	pvarg, 
						int				cArgs,
						VARIANT		*	pVarResult);
	HRESULT			DoInvoke(
						DISPID			dispid,
						WORD			wFlags,
						VARIANTARG	*	pvarg,
						int				cArgs,
						VARIANT		*	pVarResult);
	DISPID			GetDispid(
						LPCTSTR			szMethod);
private:
	IDispatch	*	m_pdisp;
};
