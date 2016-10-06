#pragma once
#include "dispobject.h"

class CSciFileOpen : public CDispObject
{
public:
	CSciFileOpen(void);
	virtual ~CSciFileOpen(void);
//	void			AddFileExtension(
//						LPCTSTR		fileExtension,
//						LPCTSTR		fileDescription);
//	BOOL			SelectDataFile(
//						HWND		hwnd,
//						LPTSTR		szDataFile,
//						UINT		nBufferSize);
	void			SetBaseWorkingDirectory(
						LPCTSTR		szBaseWorkingDirectory);
	BOOL			SelectCalibrationFile(
						HWND		hwndParent,
						LPTSTR		szCalibration,
						UINT		nBufferSize);

};
