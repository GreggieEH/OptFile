#pragma once
#include "dispobject.h"

class CCalibrationStandard	 : public CDispObject
{
public:
	CCalibrationStandard(void);
	virtual ~CCalibrationStandard(void);
	BOOL			InitNew();
	BOOL			loadFromXML(
						IDispatch	*	xml);
	BOOL			saveToXML(
						IDispatch	*	xml);
	BOOL			GetfileName(
						LPTSTR			szFileName,
						UINT			nBufferSize);
	void			SetfileName(
						LPCTSTR			szfileName);
	BOOL			GetamLoaded();
	void			SetamLoaded(
						BOOL			amLoaded);
	BOOL			Getwavelengths(
						long		*	nValues,
						double		**	ppWaves);
	BOOL			Getcalibration(
						long		*	nValues,
						double		**	ppcalibration);
	BOOL			loadExcelFile(
						LPCTSTR			szFileName);
	BOOL			checkExcelFile(
						LPCTSTR			szFileName);
protected:
	BOOL			VariantToDoubleArray(
						VARIANT			*	Value,
						long			*	nValues,
						double			**	ppValues);
};
