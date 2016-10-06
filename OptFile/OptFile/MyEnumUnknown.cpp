#include "StdAfx.h"
#include "MyEnumUnknown.h"
//#include "Utils.h"

CMyEnumUnknown::CMyEnumUnknown(IUnknown * punkHost) :
	m_pUnkHost(punkHost),
	m_cRefs(0),
	m_cUnk(0),
	m_ppUnk(NULL),
	m_cCurIndex(0)
{
}

CMyEnumUnknown::~CMyEnumUnknown(void)
{
	if (NULL != this->m_ppUnk)
	{
		for (ULONG i = 0; i<m_cUnk; i++)
		{
			Utils_RELEASE_INTERFACE(this->m_ppUnk[i]);
		}
		delete[] this->m_ppUnk;
		this->m_ppUnk = NULL;
	}
	this->m_cUnk = 0;
}

// IUnknown methods
STDMETHODIMP CMyEnumUnknown::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IEnumUnknown == riid)
	{
		*ppv = this;
		this->AddRef();
		return S_OK;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CMyEnumUnknown::AddRef()
{
	if (NULL != this->m_pUnkHost)
		this->m_pUnkHost->AddRef();
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyEnumUnknown::Release()
{
	ULONG				cRefs;
	if (NULL != this->m_pUnkHost)
		this->m_pUnkHost->Release();
	cRefs = --m_cRefs;
	if (0 == cRefs)
		delete this;
	return cRefs;
}

// IEnumUnknown methods
STDMETHODIMP CMyEnumUnknown::Next(
	ULONG			celt,
	IUnknown	**	rgelt,
	ULONG		*	pceltFetched)
{
	ULONG					i;

	// check the input
	if (1 != celt && NULL == pceltFetched) return E_POINTER;
	if (NULL == rgelt) return E_POINTER;
	// check the current index
	if (this->m_cCurIndex >= this->m_cUnk)
	{
		if (NULL != pceltFetched) *pceltFetched = 0;
		return S_FALSE;
	}
	i = 0;
	while (i < celt && this->m_cCurIndex < this->m_cUnk)
	{
		rgelt[i] = this->m_ppUnk[this->m_cCurIndex];
		this->m_ppUnk[this->m_cCurIndex]->AddRef();
		// increment the indices
		i++;
		this->m_cCurIndex++;
	}
	return i == celt ? S_OK : S_FALSE;
}

STDMETHODIMP CMyEnumUnknown::Skip(
	ULONG			celt)
{
	if ((celt + this->m_cCurIndex) >= this->m_cUnk)
	{
		return S_FALSE;
	}
	this->m_cCurIndex += celt;
	return S_OK;
}

STDMETHODIMP CMyEnumUnknown::Reset(void)
{
	this->m_cCurIndex = 0;
	return S_OK;
}

STDMETHODIMP CMyEnumUnknown::Clone(
	IEnumUnknown ** ppenum)
{
	CMyEnumUnknown		*	pMyEnum;
	HRESULT					hr;

	*ppenum = NULL;
	pMyEnum = new CMyEnumUnknown(this->m_pUnkHost);
	if (NULL != pMyEnum)
	{
		hr = pMyEnum->Init(this->m_cUnk, this->m_ppUnk, this->m_cCurIndex);
		if (SUCCEEDED(hr)) hr = pMyEnum->QueryInterface(IID_IEnumUnknown, (LPVOID*)ppenum);
		if (FAILED(hr)) delete pMyEnum;
	}
	else hr = E_OUTOFMEMORY;
	return hr;
}

// initialization
HRESULT CMyEnumUnknown::Init(
	ULONG				cUnk,
	IUnknown		**	ppUnk,
	ULONG				cStartIndex)
{
	ULONG					i;

	this->m_ppUnk = new IUnknown *[cUnk];
	if (NULL == this->m_ppUnk) return E_OUTOFMEMORY;
	for (i = 0; i<cUnk; i++)
	{
		this->m_ppUnk[i] = ppUnk[i];
		this->m_ppUnk[i]->AddRef();
	}
	this->m_cUnk = cUnk;
	this->m_cCurIndex = cStartIndex;
	return S_OK;
}
