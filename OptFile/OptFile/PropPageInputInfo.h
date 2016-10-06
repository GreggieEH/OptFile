#pragma once
#include "myproppage.h"

class CMyObject;

class CPropPageInputInfo : public CMyPropPage
{
public:
	CPropPageInputInfo(CMyObject * pMyObject);
	virtual ~CPropPageInputInfo(void);
protected:
	virtual BOOL	OnInitDialog();
	virtual BOOL	OnCommand(
						WORD			wmID,
						WORD			wmEvent);
	virtual void	OnEditReturnClicked(
						UINT			nID);
	BOOL			GetInputInfo(
						IDispatch	**	ppdisp);
	void			DisplayLampDistance();
	void			SetLampDistance(
						double			lampDistance);
	void			UpdateLampDistance();
	void			DisplayAvailableConfigs();
	void			DisplayInputConfig();
	void			SetInputConfig(
						LPCTSTR			szInputConfig);
	void			DisplayComment();
	void			OnClickedChkUserSet();
};
