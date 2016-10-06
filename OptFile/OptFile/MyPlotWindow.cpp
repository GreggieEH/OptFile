#include "stdafx.h"
#include "MyPlotWindow.h"

CMyPlotWindow::CMyPlotWindow(HWND hwnd) :
	m_hwndPlot(hwnd),
	m_nValues(0),
	m_pX(NULL),
	m_pY(NULL),
	m_xRat(0.0),
	m_yRat(0.0),
	m_startX(0.0),
	m_startY(0.0)
{
	this->m_ptTopLeft.x	= 0;
	this->m_ptTopLeft.y = 0;
}

CMyPlotWindow::~CMyPlotWindow(void)
{
	this->ClearPlotData();
}

void CMyPlotWindow::SetPlotData(
						long		nValues,
						double	*	pX,
						double	*	pY)
{
	this->ClearPlotData();
	long			i;
	this->m_nValues	= nValues;
	this->m_pX	= new double [this->m_nValues];
	this->m_pY	= new double [this->m_nValues];
	for (i=0; i<this->m_nValues; i++)
	{
		this->m_pX[i]	= pX[i];
		this->m_pY[i]	= pY[i];
	}
}

void CMyPlotWindow::OnPaint()
{
	RECT			rc;
	HDC				hdc;
	PAINTSTRUCT		ps;
	HDC				hdcMem;
	HBITMAP			hbmMem;

	GetClientRect(this->m_hwndPlot, &rc);
	hdc = BeginPaint(this->m_hwndPlot, &ps);
	hdcMem	= CreateCompatibleDC(hdc);
	hbmMem	= CreateCompatibleBitmap(hdc, rc.right - rc.left, rc.bottom - rc.top);
	SelectObject(hdcMem, (HGDIOBJ) hbmMem);
	this->DoPaint(hdcMem, &rc);
	BitBlt(hdc, 0, 0, rc.right - rc.left, rc.bottom - rc.top, hdcMem, 0, 0, SRCCOPY);
	DeleteObject((HGDIOBJ) hbmMem);
	DeleteDC(hdcMem);
	EndPaint(this->m_hwndPlot, &ps);
}

void CMyPlotWindow::OnMouseMove(
						long		xPos,
						long		yPos,
						LPTSTR		szString,
						UINT		nBufferSize)
{
	if (0.0 != this->m_xRat || 0.0 != this->m_yRat)
	{
		double		x	= this->m_startX + ((xPos - this->m_ptTopLeft.x) * this->m_xRat);
		double		y	= this->m_startY - ((yPos - this->m_ptTopLeft.y) * this->m_yRat);
		_stprintf_s(szString, nBufferSize, TEXT("X = %4.1f  Y = %8.4g"), x, y);
	}
	else
	{
		StringCchPrintf(szString, nBufferSize, TEXT("Not Set"));
	}
}

void CMyPlotWindow::ClearPlotData()
{
	if (NULL != this->m_pX)
	{
		delete [] this->m_pX;
		this->m_pX	= NULL;
	}
	if (NULL != this->m_pY)
	{
		delete [] this->m_pY;
		this->m_pY	= NULL;
	}
	this->m_nValues	= 0;
}

void CMyPlotWindow::DoPaint(
						HDC			hdc,
						LPCRECT		prc)
{
	HBRUSH			hbrBackground;		// background brush
	double			xRange[2];
	double			yRange[2];
	long			i;
	HPEN			hpen;
	HPEN			hOldPen;
	long			x, y;

	this->m_xRat	= 0.0;
	this->m_yRat	= 0.0;
	if (IsRectEmpty(prc)) return;

	// paint background white
	hbrBackground	= CreateSolidBrush(RGB(255, 255, 255));
	FillRect(hdc, prc, hbrBackground);
	DeleteObject((HGDIOBJ) hbrBackground);
	// get the data range
	if (NULL == this->m_pX || NULL == this->m_pY)
	{
		return;
	}
	xRange[0]	= this->m_pX[0];
	xRange[1]	= xRange[0];
	yRange[0]	= this->m_pY[0];
	yRange[1]	= yRange[0];
	for (i=1; i<this->m_nValues; i++)
	{
		xRange[0]	= this->m_pX[i] < xRange[0] ? this->m_pX[i] : xRange[0];
		xRange[1]	= this->m_pX[i] > xRange[1] ? this->m_pX[i] : xRange[1];
		yRange[0]	= this->m_pY[i] < yRange[0] ? this->m_pY[i] : yRange[0];
		yRange[1]	= this->m_pY[i] > yRange[1] ? this->m_pY[i] : yRange[1];
	}
	if (xRange[1] > xRange[0] && yRange[1] > yRange[0])
	{
		this->m_xRat	= (xRange[1] - xRange[0]) / (prc->right - prc->left);
		this->m_yRat	= (yRange[1] - yRange[0]) / (prc->bottom - prc->top);
		this->m_ptTopLeft.x	= prc->left;
		this->m_ptTopLeft.y = prc->bottom;
		this->m_startX	= xRange[0];
		this->m_startY	= yRange[0];
		// create pen
		hpen		= CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
		hOldPen		= (HPEN) SelectObject(hdc, (HGDIOBJ) hpen);
		// draw curve
		x			= this->m_ptTopLeft.x + (long) floor((this->m_pX[0] - this->m_startX) / this->m_xRat);
		y			= this->m_ptTopLeft.y - (long) floor((this->m_pY[0] - this->m_startY) / this->m_yRat);
		MoveToEx(hdc, x, y, NULL);
		for (i=1; i<this->m_nValues; i++)
		{
			x		= this->m_ptTopLeft.x + (long) floor((this->m_pX[i] - this->m_startX) / this->m_xRat);
			y		= this->m_ptTopLeft.y - (long) floor((this->m_pY[i] - this->m_startY) / this->m_yRat);
			LineTo(hdc, x, y);
		}
		// free pen
		SelectObject(hdc, (HGDIOBJ) hOldPen);
		DeleteObject((HGDIOBJ) hpen);
	}
}
