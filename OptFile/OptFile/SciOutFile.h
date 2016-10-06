#pragma once
#include "DispObject.h"
class CSciOutFile : public CDispObject
{
public:
	CSciOutFile();
	virtual ~CSciOutFile();
	void			SetfileName(
						LPCTSTR			szfileName);
	void			loadFile();
	BOOL			GetxValues(
						long		*	nX,
						double		**	ppX);
	BOOL			GetyValues(
						long		*	nX,
						double		**	ppY);
};

