#pragma once

class CMyObject;

class CMyPropPage
{
public:
	CMyPropPage(CMyObject * pMyObject, UINT nID, UINT nTitleID);
	virtual ~CMyPropPage(void);
	virtual HPROPSHEETPAGE	doCreatePropPage();
	HWND			GetMyPage();
protected:
	virtual BOOL	PropPageProc(
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam);
	virtual BOOL	OnInitDialog();
	virtual BOOL	OnCommand(
						WORD			wmID,
						WORD			wmEvent);
	virtual BOOL	OnNotify(
						LPNMHDR			pnmh);
	void			SubclassEditBox(
						UINT			nID);
	void			SubclassControl(
						UINT			nID);
	virtual void	OnEditReturnClicked(
						UINT			nID);
	virtual LRESULT	SubclassProc(
						HWND			hwnd,
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam);
	BOOL			GetOurObject(
						IDispatch	**	ppdisp);
	void			SetPageChanged();
private:
	CMyObject	*	m_pMyObject;
	HWND			m_hwndDlg;			// dialog handle
	UINT			m_nID;				// property page id
	UINT			m_nTitleID;			// ID string for title

	// structure sent in with the subclass procedures
	struct SUBCLASS_STRUCT
	{
		CMyPropPage	*	pMyPropPage;
		UINT			nID;			// control ID
		HWND			hwnd;			// control window
		WNDPROC			wpOrig;			// original window procedure
	};

// property page dialog procedure
friend LRESULT CALLBACK	MyPropPage(HWND, UINT, WPARAM, LPARAM);
// subclass procedures
friend LRESULT CALLBACK MyEditSubclass(HWND, UINT, WPARAM, LPARAM);
friend LRESULT CALLBACK MySubclassProc(HWND, UINT, WPARAM, LPARAM);
};
