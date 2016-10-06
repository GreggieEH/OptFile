#pragma once

class CMyPlotWindow;

class CDlgDisplayRadiance
{
public:
	CDlgDisplayRadiance(void);
	~CDlgDisplayRadiance(void);
	void			SetArrayData(
						long			nValues,
						double		*	pX,
						double		*	pY);
	void			DoOpenDialog(
						HWND			hwndParent,
						HINSTANCE		hInst);
protected:
	BOOL			DlgProc(
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam);
	BOOL			OnInitDialog();
	BOOL			OnCommand(
						WORD			wmID,
						WORD			wmEvent);
	void			SubclassControl(
						UINT			nID);
	LRESULT			SubclassProc(
						HWND			hwnd,
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam);

private:
	HWND			m_hwndDlg;
	CMyPlotWindow*	m_pMyPlotWindow;
	long			m_nValues;
	double		*	m_pX;
	double		*	m_pY;

friend LRESULT CALLBACK	DlgProcDisplayRadiance(HWND, UINT, WPARAM, LPARAM);
	// structure sent in with the subclass procedures
	struct SUBCLASS_STRUCT
	{
		CDlgDisplayRadiance*	pDlg;
		UINT					nID;			// control ID
		HWND					hwnd;			// control window
		WNDPROC					wpOrig;			// original window procedure
	};

friend LRESULT CALLBACK SubclassProcPlotWindow(HWND, UINT, WPARAM, LPARAM);

};
