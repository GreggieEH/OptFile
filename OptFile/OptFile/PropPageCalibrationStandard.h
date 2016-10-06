#pragma once
#include "myproppage.h"
#include "PlotWindowEventHandler.h"

class CMyObject;
//class CMyPlotWindow;
class CPlotWindow;

class CPropPageCalibrationStandard : public CMyPropPage,
									 public CPlotWindowEventHandler
{
public:
	CPropPageCalibrationStandard(CMyObject * pMyObject);
	virtual ~CPropPageCalibrationStandard(void);
	// plot window events
	void			OnPlotWindowMouseMove(
						double			x,
						double			y);


protected:
	virtual BOOL	OnInitDialog();
	virtual BOOL	OnCommand(
						WORD			wmID,
						WORD			wmEvent);
	virtual BOOL	OnNotify(
						LPNMHDR			pnmh);
//	virtual LRESULT	SubclassProc(
//						HWND			hwnd,
//						UINT			uMsg,
//						WPARAM			wParam,
//						LPARAM			lParam);
	BOOL			GetCalibrationStandard(
						IDispatch	**	ppdisp);
	BOOL			GetFileName(
						LPTSTR			szFileName,
						UINT			nBufferSize);
	void			SetFileName(
						LPCTSTR			szFileName);
	BOOL			GetAmLoaded();
	void			SetAmLoaded(
						BOOL			fAmLoaded);
	BOOL			Getwavelengths(
						long		*	nValues,
						double		**	ppWaves);
	BOOL			Getcalibration(
						long		*	nValues,
						double		**	ppcalibration);
	BOOL			SelectCalibrationFile();
	void			DisplayCurrentData();
	void			DisplayexpectedDistance();
	void			UpdateexpectedDistance();
	void			DisplaysourceDistance();
	void			UpdatesourceDistance();
private:
//	CMyPlotWindow*	m_pMyPlotWindow;
	CPlotWindow	*	m_pMyPlotWindow;
	BOOL			m_fAllowClose;
};
