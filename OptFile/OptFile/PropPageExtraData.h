#pragma once
#include "MyPropPage.h"

class CMyObject;
class CMyPlotWindow;

class CPropPageExtraData : public CMyPropPage
{
public:
	CPropPageExtraData(CMyObject * pMyObject);
	virtual ~CPropPageExtraData();
//	virtual HPROPSHEETPAGE	doCreatePropPage();
protected:
	virtual BOOL	OnInitDialog();
	virtual LRESULT	SubclassProc(
						HWND			hwnd,
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam);
	virtual BOOL	OnNotify(
						LPNMHDR			pnmh);
	void			DisplayExtraData();
	BOOL			GetTitleString(
						LPTSTR			szTitle,
						UINT			nBufferSize);
	void			GetExtraDataTitles();
private:
	CMyPlotWindow*	m_pMyPlotWindow;
	ULONG			m_nExtraData;
	LPTSTR		*	m_pszExtraData;
	ULONG			m_nCurrentData;
};

