#pragma once
#include "myproppage.h"
#include "PlotWindowEventHandler.h"

class CMyObject;
//class CMyPlotWindow;
class CPlotWindow;

class CPropPageGratingScanInfo : public CMyPropPage,
								 public CPlotWindowEventHandler
{
public:
	CPropPageGratingScanInfo(CMyObject * pMyObject);
	virtual ~CPropPageGratingScanInfo(void);
	virtual void	OnPlotWindowMouseMove(
						double			x,
						double			y);
protected:
	virtual BOOL	OnInitDialog();
	virtual BOOL	OnNotify(
						LPNMHDR			pnmh);
	virtual void	OnEditReturnClicked(
						UINT			nID);
//	virtual LRESULT	SubclassProc(
//						HWND			hwnd,
//						UINT			uMsg,
//						WPARAM			wParam,
//						LPARAM			lParam);
	void			DisplayGratingScan(
						long			scanIndex);
	long			GetNumGratingScans();
	BOOL			GetGratingScan(
						long			index,
						IDispatch	**	ppdispGratingScan);
	void			DisplayTimeStamp();
	double			GetOutputSlitWidth();
private:
//	CMyPlotWindow*	m_pMyPlotWindow;
	CPlotWindow	*	m_pMyPlotWindow;
};
