#include "stdafx.h"
#include "DynamicArray.h"


CDynamicArray::CDynamicArray() :
	m_nAllocate(100),			// default number of values to allocate
	m_nValues(0),			// number of valid values
	m_pX(NULL),					// x array
	m_pY(NULL)					// y array
{
	this->AllocateArrays(this->m_nAllocate);
}

CDynamicArray::~CDynamicArray()
{
	this->AllocateArrays(0);
}

void CDynamicArray::AllocateArrays(
	long			nAlloc)
{
	if (NULL != this->m_pX)
	{
		delete[] this->m_pX;
		this->m_pX = NULL;
	}
	if (NULL != this->m_pY)
	{
		delete[] this->m_pY;
		this->m_pY = NULL;
	}
	this->m_nValues = 0;
	this->m_nAllocate = nAlloc;
	if (this->m_nAllocate > 0)
	{
		this->m_pX = new double[this->m_nAllocate];
		this->m_pY = new double[this->m_nAllocate];
	}
}

void CDynamicArray::AddValues(
	double			x,
	double			y)
{
	if (this->m_nValues < this->m_nAllocate)
	{
		// add the new value
		this->m_pX[this->m_nValues] = x;
		this->m_pY[this->m_nValues] = y;
		this->m_nValues++;
	}
	else
	{
		// need to allocate more space
		long			nAlloc = this->m_nAllocate + 100;
		double		*	tX = new double[nAlloc];
		double		*	tY = new double[nAlloc];
		long			i;
		// copy the current values
		for (i = 0; i < this->m_nValues; i++)
		{
			tX[i] = this->m_pX[i];
			tY[i] = this->m_pY[i];
		}
		// update the arrays
		delete[] this->m_pX;
		delete[] this->m_pY;
		// store the new 
		this->m_pX = tX;
		this->m_pY = tY;
		this->m_nAllocate = nAlloc;
		// add the new value
		this->m_pX[this->m_nValues] = x;
		this->m_pY[this->m_nValues] = y;
		this->m_nValues++;
	}
}

BOOL CDynamicArray::GetXArray(
	long		*	nValues,
	double		**	ppX)
{
	long		i;
	if (NULL != this->m_pX && this->m_nValues > 0)
	{
		*nValues = this->m_nValues;
		*ppX = new double[*nValues];
		for (i = 0; i < (*nValues); i++)
			(*ppX)[i] = this->m_pX[i];
		return TRUE;
	}
	else
	{
		*nValues = 0;
		*ppX = NULL;
		return FALSE;
	}
}

BOOL CDynamicArray::GetYArray(
	long		*	nValues,
	double		**	ppY)
{
	long		i;
	if (NULL != this->m_pY && this->m_nValues > 0)
	{
		*nValues = this->m_nValues;
		*ppY = new double[*nValues];
		for (i = 0; i < (*nValues); i++)
			(*ppY)[i] = this->m_pY[i];
		return TRUE;
	}
	else
	{
		*nValues = 0;
		*ppY = NULL;
		return FALSE;
	}
}
