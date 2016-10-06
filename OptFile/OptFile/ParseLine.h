#pragma once
#include "DispObject.h"
class CParseLine :
	public CDispObject
{
public:
	CParseLine();
	virtual ~CParseLine();
	void			doParseLine(
						LPCTSTR			oneLine);
	long			GetnumSubStrings();
	BOOL			getSubString(
						long			index,
						LPTSTR			szString,
						UINT			nBufferSize);
};

