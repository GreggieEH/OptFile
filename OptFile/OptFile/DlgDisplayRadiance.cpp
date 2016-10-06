#include "stdafx.h"
#include "DlgDisplayRadiance.h"
#include "MyPlotWindow.h"

CDlgDisplayRadiance::CDlgDisplayRadiance(void) :
	m_hwndDlg(NULL),
	m_pMyPlotWindow(NULL),
	m_nValues(0),
	m_pX(NULL),
	m_pY(NULL)
{
}

CDlgDisplayRadiance::~CDlgDisplayRadiance(void)
{
	if (NULL != this->m_pX)
	{
		delete [] this->m_pX;
		this->m_pX = NULL;
	}
	if (NULL != this->m_pY)
	{
		delete [] this->m_pY;
		this->m_pY = NULL;
	}
	this->m_nValues = 0;
	Utils_DELETE_POINTER(this->m_pMyPlotWindow);
}

void CDlgDisplayRadiance::SetArrayData(
						long			nValues,
						double		*	pX,
						double		*	pY)
{
	if (NULL != this->m_pX)
	{
		delete [] this->m_pX;
		this->m_pX = NULL;
	}
	if (NULL != this->m_pY)
	{
		delete [] this->m_pY;
		this->m_pY = NULL;
	}
	this->m_nValues = 0;
	if (nValues > 0 && NULL != pX && NULL != pY)
	{
		long			i;
		this->m_nValues	= nValues;
		this->m_pX		= new double [this->m_nValues];
		this->m_pY		= new double [this->m_nValues];
		for (i=0; i<this->m_nValues; i++)
		{
			this->m_pX[i]	= pX[i];
			this->m_pY[i]	= pY[i];
		}
	}
}

void CDlgDisplayRadiance::DoOpenDialog(
						HWND			hwndParent,
						HINSTANCE		hInst)
{
	DialogBoxParam(hInst, MAKEINTRESOURCE(IID_DIALOGDisplayRadiance), hwndParent,
		(DLGPROC) DlgProcDisplayRadiance, (LPARAM) this);
}

LRESULT CALLBACK DlgProcDisplayRadiance(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CDlgDisplayRadiance	*	pDlg	= NULL;
	if (WM_INITDIALOG == uMsg)
	{
		pDlg = (CDlgDisplayRadiance*) lParam;
		SetWindowLongPtr(hwndDlg, GWLP_USERDATA, lParam);
		pDlg->m_hwndDlg	= hwndDlg;
	}
	else
	{
		pDlg = (CDlgDisplayRadiance*) GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
	}
	if (NULL != pDlg)
		return pDlg->DlgProc(uMsg, wParam, lParam);
	else
		return FALSE;
}

BOOL CDlgDisplayRadiance::DlgProc(
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
	default:
		break;
	}
	return FALSE;
}

BOOL CDlgDisplayRadiance::OnInitDialog()
{
	// subclass the plot window
	this->SubclassControl(IDC_PLOT);
	// create the plot window object
	this->m_pMyPlotWindow		= new CMyPlotWindow(GetDlgItem(this->m_hwndDlg, IDC_PLOT));
	// set the data
	this->m_pMyPlotWindow->SetPlotData(this->m_nValues, this->m_pX, this->m_pY);
	return TRUE;
}

BOOL CDlgDisplayRadiance::OnCommand(
						WORD			wmID,
						WORD			wmEvent)
{
	switch (wmID)
	{
	case IDOK:
	case IDCANCEL:
		EndDialog(this->m_hwndDlg, wmID);
		return TRUE;
	default:
		break;
	}
	return FALSE;
}

void CDlgDisplayRadiance::SubclassControl(
						UINT			nID)
{
	CDlgDisplayRadiance::SUBCLASS_STRUCT * pSubclassStr = new CDlgDisplayRadiance::SUBCLASS_STRUCT;
	pSubclassStr->nID			= nID;
	pSubclassStr->pDlg			= this;
	pSubclassStr->hwnd			= GetDlgItem(this->m_hwndDlg, nID);
	SetWindowLongPtr(pSubclassStr->hwnd, GWLP_USERDATA, (LONG) pSubclassStr);
	pSubclassStr->wpOrig		= (WNDPROC) SetWindowLongPtr(pSubclassStr->hwnd,
								GWLP_WNDPROC, (LONG) SubclassProcPlotWindow);
}

LRESULT CDlgDisplayRadiance::SubclassProc(
						HWND			hwnd,
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam)
{
	CDlgDisplayRadiance::SUBCLASS_STRUCT*	pSubclassStruct = 
		(CDlgDisplayRadiance::SUBCLASS_STRUCT*) GetWindowLongPtr(hwnd, GWLP_USERDATA);
	UINT			nID		= GetDlgCtrlID(hwnd);
	if (IDC_PLOT == nID)
	{
		switch (uMsg)
		{
		case WM_PAINT:
			if (NULL != this->m_pMyPlotWindow)
			{
				this->m_pMyPlotWindow->OnPaint();
				return 0;
			}
			break;
		case WM_MOUSEMOVE:
			if (NULL != this->m_pMyPlotWindow)
			{
				TCHAR			szString[MAX_PATH];
				szString[0]	= '\0';
				this->m_pMyPlotWindow->OnMouseMove(
					GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), szString, MAX_PATH);
				SetDlgItemText(this->m_hwndDlg, IDC_STATUS, szString);
			}
		default:
			break;
		}
	}
	return CallWindowProc(pSubclassStruct->wpOrig, hwnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK SubclassProcPlotWindow(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CDlgDisplayRadiance::SUBCLASS_STRUCT * pSubclassStr	= 
		(CDlgDisplayRadiance::SUBCLASS_STRUCT * ) GetWindowLongPtr(hwnd, GWLP_USERDATA);
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
	return pSubclassStr->pDlg->SubclassProc(hwnd, uMsg, wParam, lParam);
}


