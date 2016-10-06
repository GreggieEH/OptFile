// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>
#include <Windowsx.h>
#include <ole2.h>
#include <olectl.h>

#include <shlobj.h>
#include <shlwapi.h>
#include <stdio.h>
#include <tchar.h>
#include <math.h>
#include <Propvarutil.h>

#include <vector>

#include <strsafe.h>
#include "G:\Users\Greg\Documents\Visual Studio 2015\Projects\MyUtils\MyUtils\myutils.h"
#include "resource.h"

class CServer;
CServer * GetServer();

// translate array from JScript code
BOOL TranslateJScriptArray(
	VARIANT		*	pVarResult,
	long		*	nValues,
	double		**	ppValues);
// create array for JScript code
BOOL CreateJScriptArray(
	long			nValues,
	double		*	pValues,
	VARIANT		*	pVarg);
// get application window
HWND GetApplicationWindow();

// definitions
#define				MY_CLSID					CLSID_OPTFile
#define				MY_LIBID					LIBID_OPTFile
#define				MY_IID						IID_IOPTFile
#define				MY_IIDSINK					IID__OPTFile

#define				MAX_CONN_PTS				2
#define				CONN_PT_CUSTOMSINK			0
#define				CONN_PT__clsIDataSet		1

#define				FRIENDLY_NAME				TEXT("OPTFile")
#define				PROGID						TEXT("Sciencetech.OPTFile.1")
#define				VERSIONINDEPENDENTPROGID	TEXT("Sciencetech.OPTFile")

// root node name
#define				ROOT_NODE_NAME				TEXT("OPTFile")


// object names
#define				PLOT_WINDOW					L"SciPlotWindow"

// file open displayed curve
#define				OUR_CURVE					L"OPTCurve"
