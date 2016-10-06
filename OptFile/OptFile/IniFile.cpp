// IniFile.cpp: implementation of the CIniFile class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "IniFile.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CIniFile::CIniFile() :
	m_szIniFile(NULL)			// initialization file
{
}

CIniFile::~CIniFile()
{
	if (m_szIniFile)
	{
		delete [] m_szIniFile;
		m_szIniFile = NULL;
	}
}

BOOL CIniFile::Init(
							LPCTSTR				szIniFile)
{
	// initialization file
	int				slen = 0;
	if (szIniFile)
	{
		slen = lstrlen(szIniFile);
	}
	if (slen > 0)
	{
		m_szIniFile = new TCHAR [slen+1];
		lstrcpy(m_szIniFile, szIniFile);
		return TRUE;
	}
	return FALSE;
}

BOOL CIniFile::ReadIntValue(
					LPCTSTR				szSectionName,
					LPCTSTR				szValueName,
					long				defValue,
					long			*	pvalue)
{
	BOOL			fSuccess = FALSE;		// success flag
	TCHAR			szValue[40];			// value string
	TCHAR			szDefault[40];			// default value string
//	WCHAR			wszValue[40];
	BSTR			bstr;					// BSTR
	HRESULT			hr;
	TCHAR			szString[40];
	int				slen;					// string length
	int				i;						// index over the string
	int				c;						// string index
	BOOL			fDone;					// completed flag
	TCHAR			chBlank = TEXT(" ")[0];
	TCHAR			chNeg = TEXT("-")[0];
	TCHAR			chPos = TEXT("+")[0];

	// format the default value
	hr = VarBstrFromI4(defValue, LOCALE_SYSTEM_DEFAULT, 0, &bstr);
	if (SUCCEEDED(hr))
	{
		StringCchCopy(szDefault, 40, bstr);
		SysFreeString(bstr);
	}
	// call the read string function
	if (SUCCEEDED(hr))
	{
		fSuccess = ReadStringValue(szSectionName, szValueName,
			szDefault, szValue, 40);
		// only take the first numeric characters
		slen = lstrlen(szValue);
		fDone = FALSE;
		i = 0;
		c = 0;
		while (i < slen && !fDone)
		{
			if (chBlank != szValue[i])
			{
				fDone =	(chNeg != szValue[i])			&&
						(chPos != szValue[i])			&&
						!(IsCharAlphaNumeric(szValue[i]) && !IsCharAlpha(szValue[i]));
				if (!fDone)
				{
					szString[c++] = szValue[i];
					i++;
				}
			}
			else i++;					// skip past the first blanks
		}
		szString[c] = 0;					// null terminate
	}
	// read the string
	if (fSuccess)
	{
//		MultiByteToWideChar(CP_ACP, 0, szString, -1,
//			wszValue, 40);
		hr = VarI4FromStr(szString, LOCALE_SYSTEM_DEFAULT, 0, pvalue);
		fSuccess = SUCCEEDED(hr);
	}
	return fSuccess;
}

BOOL CIniFile::ReadStringValue(
					LPCTSTR				szSectionName,
					LPCTSTR				szValueName,
					LPCTSTR				szDefault,
					LPTSTR				szOutput,
					int					nChars)
{
	DWORD			dwNRead;

	dwNRead = GetPrivateProfileString(szSectionName, szValueName,
			szDefault, szOutput, 16, m_szIniFile);
	return dwNRead > 0;
}

BOOL CIniFile::ReadDoubleValue(
					LPCTSTR				szSectionName,
					LPCTSTR				szValueName,
					double				defValue,
					double			*	pvalue)
{
	BOOL			fSuccess = FALSE;		// success flag
	TCHAR			szValue[40];			// value string
	TCHAR			szDefault[40];			// default value string
//	WCHAR			wszValue[40];
	BSTR			bstr;					// BSTR
	HRESULT			hr;
	TCHAR			szString[40];
	int				slen;					// string length
	int				i;						// index over the string
	int				c;						// string index
	BOOL			fDone;					// completed flag
	TCHAR			chBlank = TEXT(" ")[0];	// blank character
	TCHAR			chDot = TEXT(".")[0];	// decimal
	TCHAR			chNeg = TEXT("-")[0];
	TCHAR			chPos = TEXT("+")[0];

	// format the default value
	hr = VarBstrFromR8(defValue, LOCALE_SYSTEM_DEFAULT, 0, &bstr);
	if (SUCCEEDED(hr))
	{
//		WideCharToMultiByte(CP_ACP, 0, bstr, -1,
//			szDefault, 40 * sizeof(TCHAR), NULL, NULL);
		StringCchCopy(szDefault, 40, bstr);
		SysFreeString(bstr);
	}
	// call the read string function
	if (SUCCEEDED(hr))
	{
		fSuccess = ReadStringValue(szSectionName, szValueName,
			szDefault, szValue, 40);
		// clear the leading blanks and trailing non numerics
		slen = lstrlen(szValue);
		fDone = FALSE;
		i = 0;
		c = 0;
		while (i < slen && !fDone)
		{
			if (chBlank != szValue[i])
			{
				fDone = (chDot != szValue[i])			&& 
						(chNeg != szValue[i])			&&
						(chPos != szValue[i])			&&
						!(IsCharAlphaNumeric(szValue[i]) && !IsCharAlpha(szValue[i]));
				if (!fDone)
				{
					szString[c++] = szValue[i++];
				}
			}
			else i++;			// skip past blanks
		}
		// null terminate
		szString[c] = 0;
	}
	// read the string
	if (fSuccess)
	{
		float			fval;
		if (1 == _stscanf_s(szString, TEXT("%g"), &fval))
		{
			fSuccess = TRUE;
			*pvalue = fval;
		}
/*
		MultiByteToWideChar(CP_ACP, 0, szString, -1,
			wszValue, 40);
		hr = VarR8FromStr(wszValue, LOCALE_SYSTEM_DEFAULT, 0,
			pvalue);
		fSuccess = SUCCEEDED(hr);
*/
	}
	return fSuccess;
}

BOOL CIniFile::WriteIntValue(
					LPCTSTR				szSectionName,
					LPCTSTR				szValueName,
					long				value)
{
	BSTR			bstr;
	HRESULT			hr;
	TCHAR			szValue[40];
	BOOL			fSuccess = FALSE;

	// coerce the value
	hr = VarBstrFromI4(value, LOCALE_SYSTEM_DEFAULT, 0, &bstr);
	if (SUCCEEDED(hr))
	{
//		WideCharToMultiByte(CP_ACP, 0, bstr, -1, 
//			szValue, 40 * sizeof(TCHAR), NULL, NULL);
		StringCchCopy(szValue, 40, bstr);
		SysFreeString(bstr);
		fSuccess = WriteStringValue(szSectionName, szValueName, szValue);
	}
	return fSuccess;
}

BOOL CIniFile::WriteStringValue(
					LPCTSTR				szSectionName,
					LPCTSTR				szValueName,
					LPCTSTR				szValue)
{
	return WritePrivateProfileString(
			szSectionName,
			szValueName,
			szValue, 
			m_szIniFile);
}

BOOL CIniFile::WriteDoubleValue(
					LPCTSTR				szSectionName,
					LPCTSTR				szValueName,
					double				value)
{
	HRESULT				hr;
	BSTR				bstr;
	TCHAR				szValue[40];
	BOOL				fSuccess = FALSE;		// success flag

	hr = VarBstrFromR8(value, LOCALE_SYSTEM_DEFAULT, 0, &bstr);
	if (SUCCEEDED(hr))
	{
//		WideCharToMultiByte(CP_ACP, 0, bstr, -1,
//			szValue, 40 * sizeof(TCHAR), NULL, NULL);
		StringCchCopy(szValue, 40, bstr);
		SysFreeString(bstr);
		fSuccess = WriteStringValue(szSectionName, szValueName, szValue);
	}
	return fSuccess;
}
