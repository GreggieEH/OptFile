#include "stdafx.h"
#include "MyPropPage.h"
#include "MyObject.h"
#include "Server.h"

CMyPropPage::CMyPropPage(CMyObject * pMyObject, UINT nID, UINT nTitleID) :
	m_pMyObject(pMyObject),
	m_hwndDlg(NULL),			// dialog handle
	m_nID(nID),					// property page id
	m_nTitleID(nTitleID)		// ID string for title
{
}

CMyPropPage::~CMyPropPage(void)
{
}

HPROPSHEETPAGE CMyPropPage::doCreatePropPage()
{
    PROPSHEETPAGE		psp;
    
	ZeroMemory((PVOID) &psp, sizeof(PROPSHEETPAGE));
    psp.dwSize      = sizeof(PROPSHEETPAGE);
    psp.dwFlags     = PSP_USETITLE;
    psp.hInstance   = GetServer()->GetInstance();
    psp.pszTemplate = MAKEINTRESOURCE(this->m_nID);
    psp.pfnDlgProc  = (DLGPROC) MyPropPage;
    psp.pszTitle    = MAKEINTRESOURCE(this->m_nTitleID);
    psp.lParam      = (LPARAM) this;
    psp.pfnCallback = NULL;
	return CreatePropertySheetPage(&psp);
}

// property page dialog procedure
LRESULT CALLBACK MyPropPage(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CMyPropPage	*	pMyPropPage		= NULL;
	if (WM_INITDIALOG == uMsg)
	{
		LPPROPSHEETPAGE	pPropPage	= (LPPROPSHEETPAGE) lParam;
		pMyPropPage		= (CMyPropPage*) pPropPage->lParam;
		SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG) pMyPropPage);
		pMyPropPage->m_hwndDlg	= hwnd;
	}
	else
	{
		pMyPropPage		= (CMyPropPage*) GetWindowLongPtr(hwnd, GWLP_USERDATA);
	}
	if (NULL != pMyPropPage)
		return pMyPropPage->PropPageProc(uMsg, wParam, lParam);
	else
		return FALSE;
}


HWND CMyPropPage::GetMyPage()
{
	return this->m_hwndDlg;
}

BOOL CMyPropPage::PropPageProc(
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
		return this->OnInitDialog();
	case WM_COMMAND:
		return this->OnCommand(LOWORD(wParam), HIWORD(wParam));
	case WM_NOTIFY:
		return this->OnNotify((LPNMHDR) lParam);
	default:
		break;
	}
	return FALSE;
}

BOOL CMyPropPage::OnInitDialog()
{
	Utils_CenterWindow(this->GetMyPage());
	return TRUE;
}

BOOL CMyPropPage::OnCommand(
						WORD			wmID,
						WORD			wmEvent)
{
	return FALSE;
}

BOOL CMyPropPage::OnNotify(
						LPNMHDR			pnmh)
{
	return FALSE;
}

void CMyPropPage::SubclassEditBox(
						UINT			nID)
{
	CMyPropPage::SUBCLASS_STRUCT * pSubclassStr = new CMyPropPage::SUBCLASS_STRUCT;
	pSubclassStr->nID			= nID;
	pSubclassStr->pMyPropPage	= this;
	pSubclassStr->hwnd			= GetDlgItem(this->m_hwndDlg, nID);
	SetWindowLongPtr(pSubclassStr->hwnd, GWLP_USERDATA, (LONG) pSubclassStr);
	pSubclassStr->wpOrig		= (WNDPROC) SetWindowLongPtr(pSubclassStr->hwnd,
								GWLP_WNDPROC, (LONG) MyEditSubclass);
}

LRESULT CALLBACK MyEditSubclass(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CMyPropPage::SUBCLASS_STRUCT * pSubclassStr	= (CMyPropPage::SUBCLASS_STRUCT * ) GetWindowLongPtr(hwnd, GWLP_USERDATA);
	switch (uMsg)
	{
	case WM_GETDLGCODE:
        return DLGC_WANTALLKEYS; 
	case WM_CHAR:
		if (VK_RETURN == wParam)
		{
			pSubclassStr->pMyPropPage->OnEditReturnClicked(pSubclassStr->nID);
			return 0;
		}
		break;
	case WM_DESTROY:
		{
			WNDPROC			wpOrig	= pSubclassStr->wpOrig;
			delete pSubclassStr;
			SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
			SetWindowLongPtr(hwnd, GWLP_WNDPROC, (LONG) wpOrig);
			return CallWindowProc(wpOrig, hwnd, WM_DESTROY, 0, 0);
		}
	default:
		break;
	}
	return pSubclassStr->pMyPropPage->SubclassProc(hwnd, uMsg, wParam, lParam);
}


void CMyPropPage::SubclassControl(
						UINT			nID)
{
	CMyPropPage::SUBCLASS_STRUCT * pSubclassStr = new CMyPropPage::SUBCLASS_STRUCT;
	pSubclassStr->nID			= nID;
	pSubclassStr->pMyPropPage	= this;
	pSubclassStr->hwnd			= GetDlgItem(this->m_hwndDlg, nID);
	SetWindowLongPtr(pSubclassStr->hwnd, GWLP_USERDATA, (LONG) pSubclassStr);
	pSubclassStr->wpOrig		= (WNDPROC) SetWindowLongPtr(pSubclassStr->hwnd,
								GWLP_WNDPROC, (LONG) MySubclassProc);
}

LRESULT CALLBACK MySubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CMyPropPage::SUBCLASS_STRUCT * pSubclassStr	= (CMyPropPage::SUBCLASS_STRUCT * ) GetWindowLongPtr(hwnd, GWLP_USERDATA);
	switch (uMsg)
	{
	case WM_DESTROY:
		{
			WNDPROC			wpOrig	= pSubclassStr->wpOrig;
			delete pSubclassStr;
			SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
			SetWindowLongPtr(hwnd, GWLP_WNDPROC, (LONG) wpOrig);
			return CallWindowProc(wpOrig, hwnd, WM_DESTROY, 0, 0);
		}
	default:
		break;
	}
	return pSubclassStr->pMyPropPage->SubclassProc(hwnd, uMsg, wParam, lParam);
}

void CMyPropPage::OnEditReturnClicked(
						UINT			nID)
{
}

LRESULT CMyPropPage::SubclassProc(
						HWND			hwnd,
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam)
{
	CMyPropPage::SUBCLASS_STRUCT*	pSubclassStruct = (CMyPropPage::SUBCLASS_STRUCT*) GetWindowLongPtr(hwnd, GWLP_USERDATA);
	return CallWindowProc(pSubclassStruct->wpOrig, hwnd, uMsg, wParam, lParam);
}

BOOL CMyPropPage::GetOurObject(
						IDispatch	**	ppdisp)
{
	HRESULT			hr;
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*) ppdisp);
	return SUCCEEDED(hr);
}

CMyObject* CMyPropPage::GetOurObject()
{
	return this->m_pMyObject;
}


void CMyPropPage::SetPageChanged()
{
	SendMessage(GetParent(this->GetMyPage()), PSM_CHANGED, (WPARAM) this->GetMyPage(), 0);
}
