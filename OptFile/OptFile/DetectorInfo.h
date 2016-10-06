#pragma once
#include "dispobject.h"

class CDetectorInfo : public CDispObject
{
public:
	CDetectorInfo(void);
	virtual ~CDetectorInfo(void);
	BOOL			InitNew();
	BOOL			loadFromXML(
						IDispatch	*	xml);
	BOOL			saveToXML(
						IDispatch	*	xml);
	BOOL			GetamReadOnly();
	void			SetamReadOnly(
						BOOL			famReadOnly);
	BOOL			GetdetectorModel(
						LPTSTR			szDetector,
						UINT			nBufferSize);
	void			SetdetectorModel(
						LPCTSTR			szDetector);
	BOOL			GetportName(
						LPTSTR			szPortName,
						UINT			nBufferSize);
	void			SetportName(
						LPCTSTR			szPortName);
	BOOL			GetgainSetting(
						LPTSTR			szgainSetting,
						UINT			nBufferSize);
	void			SetgainSetting(
						LPCTSTR			szgainSetting);
	double			Gettemperature();
	void			Settemperature(
						double			temperature);
	BOOL			GetamDirty();
	void			clearDirty();
	long			GetnumADChannels();
	long			getADChannel(
						long			index);
	double			getLowLimit(
						long			index);
	double			getHiLimit(
						long			index);
	long			findADChannel(
						double			wavelength);
	void			addADChannel(
						long			ADChannel, 
						double			lowLimit, 
						double			hiLimit);
	void			clearADChannels();
	long			FindIndexFromChannel(
						long			channel);
	BOOL			GetADInfoString(
						LPTSTR		*	szInfo);
	void			SetADInfoString(
						LPCTSTR			szInfo);
	BOOL			GetPulseIntegrator();
};
