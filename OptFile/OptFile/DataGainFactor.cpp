#include "stdafx.h"
#include "DataGainFactor.h"

DataGainFactor::DataGainFactor() :
	CDispObject()
{
	this->doInit(L"Sci.DataGainFactor.wsc.1.00");
}

DataGainFactor::~DataGainFactor()
{
}

void DataGainFactor::SetacquireMode(
	long		acquireMode)
{
	this->setLongProperty(L"acquireMode", acquireMode);

	long		aMode = this->getLongProperty(L"acquireMode");
	BOOL		stat = FALSE;
}

void DataGainFactor::SetlockinGain(
	long		lockinGain)
{
	this->setLongProperty(L"lockinGain", lockinGain);
}

void DataGainFactor::Setdetector(
	long		detector)
{
	this->setLongProperty(L"detector", detector);
}

void DataGainFactor::SetdetectorGain(
	LPCTSTR		szdetectorGain)
{
	this->setStringProperty(L"detectorGain", szdetectorGain);
}

double DataGainFactor::GetgainFactor()
{
	return this->getDoubleProperty(L"gainFactor");
}

double DataGainFactor::GetminRawDataValue()
{
	return this->getDoubleProperty(L"minRawDataValue");
}

double DataGainFactor::GetDCOffset()
{
	return this->getDoubleProperty(L"DCOffset");
}
