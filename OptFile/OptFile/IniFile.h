// IniFile.h: interface for the CIniFile class.
//
//////////////////////////////////////////////////////////////////////
// This class manages IO with the initialization file

#if !defined(AFX_INIFILE_H__9C17CC19_BC05_478E_9350_89D65F0EB66B__INCLUDED_)
#define AFX_INIFILE_H__9C17CC19_BC05_478E_9350_89D65F0EB66B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CIniFile  
{
public:
	CIniFile();
	virtual ~CIniFile();
	BOOL		Init(
					LPCTSTR				szIniFile);		// initialization file
	BOOL		ReadIntValue(
					LPCTSTR				szSectionName,
					LPCTSTR				szValueName,
					long				defValue,
					long			*	pvalue);
	BOOL		ReadStringValue(
					LPCTSTR				szSectionName,
					LPCTSTR				szValueName,
					LPCTSTR				szDefault,
					LPTSTR				szOutput,
					int					nChars);
	BOOL		ReadDoubleValue(
					LPCTSTR				szSectionName,
					LPCTSTR				szValueName,
					double				defValue,
					double			*	pvalue);
	BOOL		WriteIntValue(
					LPCTSTR				szSectionName,
					LPCTSTR				szValueName,
					long				value);
	BOOL		WriteStringValue(
					LPCTSTR				szSectionName,
					LPCTSTR				szValueName,
					LPCTSTR				szValue);
	BOOL		WriteDoubleValue(
					LPCTSTR				szSectionName,
					LPCTSTR				szValueName,
					double				value);
private:
	LPTSTR		m_szIniFile;			// initialization file
};

#endif // !defined(AFX_INIFILE_H__9C17CC19_BC05_478E_9350_89D65F0EB66B__INCLUDED_)
