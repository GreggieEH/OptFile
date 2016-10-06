#pragma once
#include "dispobject.h"

class COutputFiles : public CDispObject
{
public:
	COutputFiles(void);
	~COutputFiles(void);
	BOOL		initNew();
	BOOL		GetfileName(
					LPTSTR			szfileName,
					UINT			nBufferSize);
	void		SetfileName(
					LPCTSTR			szfileName);
	BOOL		GettimeStamp(
					LPTSTR			sztimeStamp,
					UINT			nBufferSize);
	void		SettimeStamp(
					LPCTSTR			sztimeStamp);
	BOOL		GetsourceFile(
					LPTSTR			szsourceFile,
					UINT			nBufferSize);
	void		SetsourceFile(
					LPCTSTR			szsourceFile);
	BOOL		GetcalibrationFile(
					LPTSTR			szcalibrationFile,
					UINT			nBufferSize);
	void		SetcalibrationFile(
					LPCTSTR			szcalibrationFile);
	BOOL		GetanalysisType(
					LPTSTR			szanalysisType,
					UINT			nBufferSize);
	void		SetanalysisType(
					LPCTSTR			szanalysisType);
	BOOL		GetxValues(
					long		*	nValues,
					double		**	ppxValues);
	void		SetxValues(
					long			nValues,
					double		*	pxValues);
	BOOL		GetyValues(
					long		*	nValues,
					double		**	ppyValues);
	void		SetyValues(
					long			nValues,
					double		*	pyValues);
	long		GetarraySize();
	BOOL		GetxUnits(
					LPTSTR			szxUnits,
					UINT			nBufferSize);
	void		SetxUnits(
					LPCTSTR			szxUnits);
	BOOL		GetyUnits(
					LPTSTR			szyUnits,
					UINT			nBufferSize);
	void		SetyUnits(
					LPCTSTR			szyUnits);
	BOOL		GetamDirty();
	void		clearDirty();
	void		loadFile();
	void		saveFile();
};
