#pragma once
#include "DispObject.h"
class DataGainFactor :
	public CDispObject
{
public:
	DataGainFactor();
	virtual ~DataGainFactor();
	void		SetacquireMode(
					long		acquireMode);
	void		SetlockinGain(
					long		lockinGain);
	void		Setdetector(
					long		detector);
	void		SetdetectorGain(
					LPCTSTR		szdetectorGain);
	double		GetgainFactor();
	double		GetminRawDataValue();
	double		GetDCOffset();
};

