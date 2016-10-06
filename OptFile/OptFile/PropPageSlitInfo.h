#pragma once
#include "myproppage.h"

class CMyObject;

class CPropPageSlitInfo : public CMyPropPage
{
public:
	CPropPageSlitInfo(CMyObject * pMyObject);
	virtual ~CPropPageSlitInfo(void);
protected:
	virtual BOOL	OnInitDialog();
	virtual BOOL	OnCommand(
						WORD			wmID,
						WORD			wmEvent);
	virtual void	OnEditReturnClicked(
						UINT			nID);
	BOOL			GetSlitInfo(
						IDispatch	**	ppdisp);
	void			DisplayInputSlit();
	void			DisplayIntermediateSlit();
	void			DisplayOutputSlit();
	void			DisplaySlitInfo(
						LPCTSTR			szSlitTitle,
						UINT			nSlitName,
						UINT			nSlitWidth);
	void			SetSlitInfo(
						LPCTSTR			szSlitTitle,
						UINT			nSlitName,
						UINT			nSlitWidth);
};
