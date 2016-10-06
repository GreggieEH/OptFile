#pragma once
class CDynamicArray
{
public:
	CDynamicArray();
	~CDynamicArray();
	void			AllocateArrays(
						long			nAlloc);
	void			AddValues(
						double			x,
						double			y);
	BOOL			GetXArray(
						long		*	nValues,
						double		**	ppX);
	BOOL			GetYArray(
						long		*	nValues,
						double		**	ppY);
private:
	long			m_nAllocate;			// number of values to allocate
	long			m_nValues;				// number of valid values
	double		*	m_pX;					// x array
	double		*	m_pY;					// y array
};

