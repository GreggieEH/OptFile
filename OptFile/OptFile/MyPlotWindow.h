#pragma once

class CMyPlotWindow
{
public:
	CMyPlotWindow(HWND hwnd);
	~CMyPlotWindow(void);
	void			ClearPlotData();
	void			SetPlotData(
						long		nValues,
						double	*	pX,
						double	*	pY);
	void			OnPaint();
	void			OnMouseMove(
						long		xPos,
						long		yPos,
						LPTSTR		szString,
						UINT		nBufferSize);
protected:
	void			DoPaint(
						HDC			hdc,
						LPCRECT		prc);
private:
	HWND			m_hwndPlot;
	long			m_nValues;
	double		*	m_pX;
	double		*	m_pY;
	double			m_xRat;
	double			m_yRat;
	double			m_startX;
	double			m_startY;
	POINT			m_ptTopLeft;
};
