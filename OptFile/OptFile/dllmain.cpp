// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include "Server.h"
#include <initguid.h>
#include "MyGuids.h"

CServer * g_pServer = NULL;

BOOL WINAPI DllMain(
	HINSTANCE hinstDLL,
	DWORD fdwReason,
	LPVOID lpvReserved
)
{
	switch (fdwReason)
	{
	case DLL_PROCESS_ATTACH:
		g_pServer = new CServer(hinstDLL);
		break;
	case DLL_PROCESS_DETACH:
		Utils_DELETE_POINTER(g_pServer);
		break;
	default:
		break;
	}
	return TRUE;
}

CServer * GetServer()
{
	return g_pServer;
}

BOOL TranslateJScriptArray(
	VARIANT		*	pVarResult,
	long		*	nValues,
	double		**	ppValues)
{
	DISPID			dispid;
	long			i;
	TCHAR			szItem[10];
	long			length = 0;
	BOOL			fSuccess = FALSE;

	*nValues = 0;
	*ppValues = NULL;
	if (VT_DISPATCH == pVarResult->vt && NULL != pVarResult->pdispVal)
	{
		// get the array length
		Utils_GetMemid(pVarResult->pdispVal, TEXT("length"), &dispid);
		length = Utils_GetLongProperty(pVarResult->pdispVal, dispid);
		if (length > 0)
		{
			fSuccess = TRUE;
			*nValues = length;
			*ppValues = new double[length];
			for (i = 0; i<length; i++)
			{
				(*ppValues)[i] = 0.0;
				StringCchPrintf(szItem, 10, TEXT("%1d"), i);
				Utils_GetMemid(pVarResult->pdispVal, szItem, &dispid);
				(*ppValues)[i] = Utils_GetDoubleProperty(pVarResult->pdispVal, dispid);
			}
		}
	}
	return fSuccess;
}

// create array for JScript code
BOOL CreateJScriptArray(
	long			nValues,
	double		*	pValues,
	VARIANT		*	pVarg)
{
	long			i;
	SAFEARRAYBOUND	sab;
	VARIANT			value;
	sab.lLbound = 0;
	sab.cElements = nValues;
	VariantInit(pVarg);
	pVarg->vt = VT_ARRAY | VT_VARIANT;
	pVarg->parray = SafeArrayCreate(VT_VARIANT, 1, &sab);
	for (i = 0; i<nValues; i++)
	{
		InitVariantFromDouble(pValues[i], &value);
		SafeArrayPutElement(pVarg->parray, &i, (void*)&value);
	}
	return TRUE;
}

// forward declaration of callback function
BOOL CALLBACK ETWProc2(HWND hwnd, LPARAM lParam);

// get application window
HWND GetApplicationWindow()
{
	// get a window to use
	HWND			hwnd = NULL;
	EnumThreadWindows(GetCurrentThreadId(), ETWProc2, (LPARAM)&hwnd);
	return hwnd;
}

BOOL CALLBACK ETWProc2(HWND hwnd, LPARAM lParam)
{
	DWORD *pdw = (DWORD *)lParam;

	/*
	* If this window has no parent, then it is a toplevel
	* window for the thread.  Remember the last one we find since it
	* is probably the main window.
	*/

	if (GetParent(hwnd) == NULL)
	{
		*pdw = (DWORD)hwnd;
	}

	return TRUE;
}
