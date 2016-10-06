#include "stdafx.h"
#include "SciXML.h"

CSciXML::CSciXML(void) :
	CDispObject()
{
	this->doInit(TEXT("Sciencetech.sciXML.wsc.1.00"));
}

CSciXML::~CSciXML(void)
{
}

BOOL CSciXML::loadFromString(
						LPCTSTR			szString)
{
	HRESULT				hr;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fSuccess	= FALSE;
	size_t				slen;

	// string needs to be ANSI
	StringCchLength(szString, STRSAFE_MAX_CCH, &slen);
	if (IsTextUnicode((const void*)szString, slen, NULL))
	{
		InitVariantFromString(szString, &varg);
/*
		int nChars = WideCharToMultiByte(CP_ACP, 0, szString, -1, NULL, 0, NULL, NULL);
		char * temp = (char*)CoTaskMemAlloc(nChars + 1);
		WideCharToMultiByte(CP_ACP, 0, szString, -1, temp, nChars + 1, NULL, NULL);
		InitVariantFromString((PCWSTR)temp, &varg);
		CoTaskMemFree((LPVOID)temp);
*/
	}
	else
	{ 
		InitVariantFromString(szString, &varg);
/*
		int nChars = MultiByteToWideChar(CP_ACP, 0, (LPCCH) szString, -1, NULL, 0);
		LPTSTR temp = (LPTSTR)CoTaskMemAlloc((nChars + 1) * sizeof(TCHAR));
		MultiByteToWideChar(CP_ACP, 0, (LPCCH)szString, -1, temp, nChars + 1);
		InitVariantFromString(temp, &varg);
		CoTaskMemFree((LPVOID)temp);
*/
	}
	hr = this->InvokeMethod(TEXT("loadFromString"), &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

BOOL CSciXML::initNew()
{
	HRESULT				hr;
	VARIANT				varResult;
	BOOL				fSuccess	= FALSE;
	hr = this->InvokeMethod(TEXT("initNew"), NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

BOOL CSciXML::saveToString(
						LPTSTR		*	szString)
{
	HRESULT				hr;
	VARIANT				varResult;
	*szString	= NULL;
	hr = this->InvokeMethod(TEXT("saveToString"), NULL, 0, &varResult);
	if (SUCCEEDED(hr))
	{
		VariantToStringAlloc(varResult, szString);
		VariantClear(&varResult);
	}
	return NULL != *szString;
}

BOOL CSciXML::GetrootName(
						LPTSTR			szrootName,
						UINT			nBufferSize)
{
	return this->getStringProperty(TEXT("rootName"), szrootName, nBufferSize);
}

void CSciXML::SetrootName(
						LPCTSTR			szrootName)
{
	this->setStringProperty(TEXT("rootName"), szrootName);
}

BOOL CSciXML::GetrootNode(
						IDispatch	**	ppdisproot)
{
	HRESULT				hr;
	VARIANT				varResult;
	BOOL				fSuccess		= FALSE;
	*ppdisproot	= NULL;
	hr = this->getProperty(TEXT("rootNode"), &varResult);
	if (SUCCEEDED(hr))
	{
		if (VT_DISPATCH == varResult.vt && NULL != varResult.pdispVal)
		{
			*ppdisproot	= varResult.pdispVal;
			varResult.pdispVal->AddRef();
			fSuccess	= TRUE;
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}

long CSciXML::determineChildNodes(
						IDispatch	*	pdispnode)
{
	HRESULT				hr;
	VARIANTARG			varg;
	VARIANT				varResult;
	long				nChild		= 0;
	VariantInit(&varg);
	varg.vt			= VT_DISPATCH;
	varg.pdispVal	= pdispnode;
	varg.pdispVal->AddRef();
	hr = this->InvokeMethod(TEXT("determineChildNodes"), &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToInt32(varResult, &nChild);
	return nChild;
}

BOOL CSciXML::getChildNode(
						IDispatch	*	pdispnode,
						long			index,
						IDispatch	**	ppdispChild)
{
	HRESULT				hr;
	VARIANTARG			avarg[2];
	VARIANT				varResult;
	BOOL				fSuccess	= FALSE;
	*ppdispChild	= NULL;
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_DISPATCH;
	avarg[1].pdispVal	= pdispnode;
	pdispnode->AddRef();
	InitVariantFromInt32(index, &avarg[0]);
	hr = this->InvokeMethod(TEXT("getChildNode"), avarg, 2, &varResult);
	VariantClear(&avarg[1]);
	if (SUCCEEDED(hr))
	{
		if (VT_DISPATCH == varResult.vt && NULL != varResult.pdispVal)
		{
			*ppdispChild	= varResult.pdispVal;
			varResult.pdispVal->AddRef();
			fSuccess		= TRUE;
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}

BOOL CSciXML::getNodeName(
						IDispatch	*	pdispnode,
						LPTSTR			szNodeName,
						UINT			nBufferSize)
{
	HRESULT				hr;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fSuccess	= FALSE;
	LPTSTR				szTemp		= NULL;
	szNodeName[0]	= '\0';
	VariantInit(&varg);
	varg.vt			= VT_DISPATCH;
	varg.pdispVal	= pdispnode;
	varg.pdispVal->AddRef();
	hr = this->InvokeMethod(TEXT("getNodeName"), &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr))
	{
		VariantToString(varResult, szNodeName, nBufferSize);
		fSuccess = TRUE;
/*
		if (NULL != szTemp)
		{
			StringCchCopy(szNodeName, nBufferSize, szTemp);
			fSuccess	= TRUE;
			CoTaskMemFree((LPVOID) szTemp);
		}
*/
		VariantClear(&varResult);
	}
	return fSuccess;
}

BOOL CSciXML::getNodeValue(
						IDispatch	*	pdispnode,
						LPTSTR			szNodeValue,
						UINT			nBufferSize)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	szNodeValue[0]	= '\0';
	if (NULL == pdispnode) return FALSE;
	VariantInit(&varg);
	varg.vt			= VT_DISPATCH;
	varg.pdispVal	= pdispnode;
	varg.pdispVal->AddRef();
	hr = this->InvokeMethod(TEXT("getNodeValue"), &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr))
	{
		hr = VariantToString(varResult, szNodeValue, nBufferSize);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}

BOOL CSciXML::getNodeValue(
	IDispatch	*	pdispnode,
	LPTSTR		*	szNodeValue)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	*szNodeValue = NULL;
	if (NULL == pdispnode) return FALSE;
	VariantInit(&varg);
	varg.vt = VT_DISPATCH;
	varg.pdispVal = pdispnode;
	varg.pdispVal->AddRef();
	hr = this->InvokeMethod(TEXT("getNodeValue"), &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr))
	{
		hr = VariantToStringAlloc(varResult, szNodeValue);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}



// get the number of grating scans
long CSciXML::GetNGratingScans(
						long		*	nGratingScans,
						LPTSTR		*	arrGratingScans)
{
	TCHAR			szNode[MAX_PATH];
	IDispatch	*	pdispRoot;
	long			nChild;
	long			i;
	long			nScans		= 0;
	long			n			= 0;

	if (this->GetrootNode(&pdispRoot))
	{
		nChild = this->determineChildNodes(pdispRoot);
		for (i=0; i<nChild; i++)
		{
			if (this->GetChildNodeName(pdispRoot, i, szNode, MAX_PATH))
			{
				if (NULL != StrStrI(szNode, TEXT("SpectraData")))
				{
					if ((*nGratingScans) > nScans)
					{
//						Utils_DoCopyString(&(arrGratingScans[nScans]), szNode);
						SHStrDup(szNode, &(arrGratingScans[nScans]));
						n++;
					}
					nScans++;
				}
			}
		}
		*nGratingScans = n;
		pdispRoot->Release();
	}
	return nScans;
}

BOOL CSciXML::GetChildNodeName(
						IDispatch	*	pdispnode,
						long			index,
						LPTSTR			szChild,
						UINT			nBufferSize)
{
	IDispatch	*	pdispChild;
	BOOL			fSuccess	= FALSE;

	if (this->getChildNode(pdispnode, index, &pdispChild))
	{
		fSuccess = this->getNodeName(pdispChild, szChild, nBufferSize);
		pdispChild->Release();
	}
	return fSuccess;
}

// find named child
BOOL CSciXML::findNamedChild(
						IDispatch	*	pdispnode,
						LPCTSTR			szChild,
						IDispatch	**	ppdispchild)
{
	HRESULT			hr;
	VARIANTARG		avarg[2];
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	*ppdispchild	= NULL;
	if (NULL == pdispnode) return FALSE;
	VariantInit(&avarg[1]);
	avarg[1].vt		= VT_DISPATCH;
	avarg[1].pdispVal	= pdispnode;
	avarg[1].pdispVal->AddRef();
	InitVariantFromString(szChild, &avarg[0]);
	hr = this->InvokeMethod(TEXT("findNamedChild"), avarg, 2, &varResult);
	VariantClear(&avarg[1]);
	VariantClear(&avarg[0]);
	if (SUCCEEDED(hr))
	{
		if (VT_DISPATCH == varResult.vt && NULL != varResult.pdispVal)
		{
			*ppdispchild = varResult.pdispVal;
			varResult.pdispVal->AddRef();
			fSuccess = TRUE;
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}

// create a new node
BOOL CSciXML::createNode(
						LPCTSTR			sznodeName,
						IDispatch	**	ppdispNode)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	*ppdispNode		= NULL;
	InitVariantFromString(sznodeName, &varg);
	hr = this->InvokeMethod(TEXT("createNode"), &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr))
	{
		if (VT_DISPATCH == varResult.vt && NULL != varResult.pdispVal)
		{
			*ppdispNode = varResult.pdispVal;
			varResult.pdispVal->AddRef();
			fSuccess = TRUE;
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}

// set node value
void CSciXML::setNodeValue(
						IDispatch	*	pdispNode,
						LPCTSTR			szValue)
{
	VARIANTARG			avarg[2];
	if (NULL == pdispNode) return;
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_DISPATCH;
	avarg[1].pdispVal	= pdispNode;
	avarg[1].pdispVal->AddRef();
	InitVariantFromString(szValue, &avarg[0]);
	this->InvokeMethod(TEXT("setNodeValue"), avarg, 2, NULL);
	VariantClear(&avarg[1]);
	VariantClear(&avarg[0]);
}

// append a child node
void CSciXML::appendChildNode(
						IDispatch	*	pdispNode,
						IDispatch	*	pdispChild)
{
	VARIANTARG			avarg[2];
	if (NULL == pdispNode || NULL == pdispChild) return;
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_DISPATCH;
	avarg[1].pdispVal	= pdispNode;
	avarg[1].pdispVal->AddRef();
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_DISPATCH;
	avarg[0].pdispVal	= pdispChild;
	avarg[0].pdispVal->AddRef();
	this->InvokeMethod(TEXT("appendChildNode"), avarg, 2, NULL);
	VariantClear(&avarg[1]);
	VariantClear(&avarg[0]);
}
