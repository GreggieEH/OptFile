#include "stdafx.h"
#include "InterpolateValues.h"

CInterpolateValues::CInterpolateValues(void) :
	m_nX(0),
	m_pX(NULL),
	m_nY(0),
	m_pY(NULL)
{
}

CInterpolateValues::~CInterpolateValues(void)
{
	if (NULL != this->m_pX)
	{
		delete[] this->m_pX;
		this->m_pX = NULL;
	}
	this->m_nX = 0;
	if (NULL != this->m_pY)
	{
		delete[] this->m_pY;
		this->m_pY = NULL;
	}
	this->m_nY = 0;
}

void CInterpolateValues::SetXArray(
						long		nValues,
						double	*	pX)
{
	long		i;
	if (NULL != this->m_pX)
	{
		delete[] this->m_pX;
		this->m_pX = NULL;
	}
	this->m_nX = nValues;
	this->m_pX = new double[this->m_nX];
	for (i = 0; i < this->m_nX; i++)
	{
		this->m_pX[i] = pX[i];
	}
}

void CInterpolateValues::SetYArray(
						long		nValues,
						double	*	pY)
{
	long		i;
	if (NULL != this->m_pY)
	{
		delete[] this->m_pY;
		this->m_pY = NULL;
	}
	this->m_nY = nValues;
	this->m_pY = new double[this->m_nY];
	for (i = 0; i < this->m_nY; i++)
		this->m_pY[i] = pY[i];
}

double CInterpolateValues::interpolateValue(
						double		inputX)
{
	long		len;
	long		i;
	BOOL		fDone;
	double		x1;
	double		x2;
	double		y1;
	double		y2;

	if (NULL == this->m_pX || NULL == this->m_pY) 
	{
		return -1.0;
	}
	len = this->GetarraySize();
	if (inputX < this->m_pX[0]) {
		return this->m_pY[0];
	}
	if (inputX > this->m_pX[len - 1]) {
		return this->m_pY[len - 1];
	}
	i = 1;
	fDone = FALSE;
	while (i < len && !fDone) 
	{
		fDone = inputX < this->m_pX[i];
		if (!fDone) i++;
	}
	if (fDone) 
	{
		x1 = this->m_pX[i - 1];
		x2 = this->m_pX[i];
		y1 = this->m_pY[i - 1];
		y2 = this->m_pY[i];
		if (x2 > x1) 
		{
			return y1 + ((inputX - x1) * (y2 - y1) / (x2 - x1));
		}
		else 
			return y1;
	}
	else 
	{
		return this->m_pY[len - 1];
	}
}

long CInterpolateValues::GetarraySize()
{
	if (this->m_nX == this->m_nY)
		return this->m_nX;
	else
		return 0;
}

double CInterpolateValues::getXValue(
						long		index)
{
	if (index >= 0 && index < this->GetarraySize())
	{
		return this->m_pX[index];
	}
	else
		return -1.0;
}

double CInterpolateValues::getYValue(
						long		index)
{
	if (index >= 0 && index < this->GetarraySize())
	{
		return this->m_pY[index];
	}
	else
	{
		return -1.0;
	}
}
