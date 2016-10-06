#pragma once
// #include "dispobject.h"

class CInterpolateValues
{
public:
	CInterpolateValues(void);
	virtual ~CInterpolateValues(void);
	void			SetXArray(
						long		nValues,
						double	*	pX);
	void			SetYArray(
						long		nValues,
						double	*	pY);
	double			interpolateValue(
						double		inputX);
	long			GetarraySize();
	double			getXValue(
						long		index);
	double			getYValue(
						long		index);
private:
//	DISPID			m_dispidinterpolateValue;
	long			m_nX;
	double		*	m_pX;
	long			m_nY;
	double		*	m_pY;
};
