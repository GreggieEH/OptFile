#pragma once
#include "dispobject.h"

class CSlitInfo : public CDispObject
{
public:
	CSlitInfo(void);
	virtual ~CSlitInfo(void);
	BOOL		AmInit();
	BOOL		InitNew();
	BOOL		loadFromXML(
					IDispatch	*	xml);
	BOOL		saveToXML(
					IDispatch	*	xml);
	BOOL		GetamReadOnly();
	void		SetamReadOnly(
					BOOL			amReadOnly);
	BOOL		GetamDirty();
	void		clearDirty();
	void		addSlit(
					LPCTSTR			SlitTitle, 
					LPCTSTR			Name, 
					double			width);
	BOOL		getSlitName(
					LPCTSTR			SlitTitle,
					LPTSTR			szName,
					UINT			nBufferSize);
	void		setSlitName(
					LPCTSTR			SlitTitle,
					LPCTSTR			szName);
	double		getSlitWidth(
					LPCTSTR			SlitTitle);
	void		setSlitWidth(
					LPCTSTR			SlitTitle,
					double			slitWidth);
	BOOL		haveSlit(
					LPCTSTR			SlitTitle);
	BOOL		GetOutputSlitWidth(
					double		*	slitWidth);
};
