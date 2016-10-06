#include "StdAfx.h"
#include "MyFactory.h"
#include "Server.h"
#include "MyObject.h"

CMyFactory::CMyFactory(void) :
	m_cRefs(0)
{
}

CMyFactory::~CMyFactory(void)
{
}

// IUnknown methods
STDMETHODIMP CMyFactory::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IClassFactory == riid)
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

STDMETHODIMP_(ULONG) CMyFactory::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyFactory::Release()
{
	ULONG				cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
	{
		this->m_cRefs++;
		GetServer()->ObjectsDown();
		delete this;
	}
	return cRefs;
}

// IClassFactory methods
STDMETHODIMP CMyFactory::CreateInstance(
									IUnknown	*	pUnkOuter,
									REFIID			riid,
									void		**	ppv)
{
	// check the input parameters
	HRESULT				hr;
	CMyObject		*	pCob;		// object

	if (NULL != pUnkOuter && riid != IID_IUnknown)
		hr = CLASS_E_NOAGGREGATION;
	else
	{
		// let's create the object
		pCob = new CMyObject(pUnkOuter);
		if (pCob != NULL)
		{
			// increment the object count
			GetServer()->ObjectsUp();
			// initialize the object
			hr = pCob->Init();
			if (SUCCEEDED(hr))
			{
				hr = pCob->QueryInterface(riid, ppv);
			}
			if (FAILED(hr))
			{
				GetServer()->ObjectsDown();
				delete pCob;
			}
		}
		else
			hr = E_OUTOFMEMORY;
	}
	return hr;
}

STDMETHODIMP CMyFactory::LockServer(
									BOOL			fLock)
{
	if (fLock)
		GetServer()->LocksUp();
	else
		GetServer()->LocksDown();
	return S_OK;
}