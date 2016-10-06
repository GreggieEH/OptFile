#pragma once
class CMyEnumUnknown : public IEnumUnknown
{
public:
	CMyEnumUnknown(IUnknown * punkHost);
	~CMyEnumUnknown(void);
	// IUnknown methods
	STDMETHODIMP				QueryInterface(
		REFIID			riid,
		LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)		AddRef();
	STDMETHODIMP_(ULONG)		Release();
	// IEnumUnknown methods
	STDMETHODIMP				Next(
		ULONG			celt,
		IUnknown	**	rgelt,
		ULONG		*	pceltFetched);
	STDMETHODIMP				Skip(
		ULONG			celt);
	STDMETHODIMP				Reset(void);
	STDMETHODIMP				Clone(
		IEnumUnknown ** ppenum);
	// initialization
	HRESULT						Init(
		ULONG			cUnk,
		IUnknown	**	ppUnk,
		ULONG			cStartIndex);
private:
	IUnknown				*	m_pUnkHost;
	ULONG						m_cRefs;
	ULONG						m_cUnk;
	IUnknown				**	m_ppUnk;
	ULONG						m_cCurIndex;
};
