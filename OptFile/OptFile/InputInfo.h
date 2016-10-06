#pragma once
#include "dispobject.h"

class CInputInfo : public CDispObject
{
public:
	CInputInfo(void);
	virtual ~CInputInfo(void);
	BOOL			InitNew();
	BOOL			loadFromXML(
						IDispatch	*	xml);
	BOOL			saveToXML(
						IDispatch	*	xml);
	BOOL			GetamReadOnly();
	void			SetamReadOnly(
						BOOL			amReadOnly);
	double			GetlampDistance();
	void			SetlampDistance(
						double			lampDistance);
	BOOL			GetinputOpticConfig(
						LPTSTR			szinputOpticConfig,
						UINT			nBufferSize);
	void			SetinputOpticConfig(
						LPCTSTR			szinputOpticConfig);
	long			GetnumInputConfigs();
	BOOL			getInputConfig(
						long			index,
						LPTSTR			szInputConfig,
						UINT			nBufferSize);
	BOOL			GetamDirty();
	void			clearDirty();
	BOOL			getRadianceAvailable(
						double		*	FOV);
	BOOL			GetuserSetinput(
						double		*	userSetInput);
	void			SetuserSetinput(
						double			userSetinput);
	void			clearUserSetInput();
};
