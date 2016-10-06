#pragma once
#include "myproppage.h"

class CMyObject;
class CIniFile;
class CDetectorInfo;

class CPropPageDetectorInfo : public CMyPropPage
{
public:
	CPropPageDetectorInfo(CMyObject * pMyObject);
	virtual ~CPropPageDetectorInfo(void);
protected:
	virtual BOOL	OnInitDialog();
	virtual BOOL	OnCommand(
						WORD			wmID,
						WORD			wmEvent);
	virtual void	OnEditReturnClicked(
						UINT			nID);
	virtual BOOL	OnNotify(
						LPNMHDR			pnmh);
	void			OnDeltaPosDetectorName(
						long			iDelta);
	void			OnDeltaPosGainSetting(
						long			iDelta);
	void			OnDeltaPosChannelDisplay(
						long			iDelta);
	BOOL			GetDetectorInfo(
						IDispatch	**	ppdisp);
	BOOL			GetDetectorInfo(
						CDetectorInfo*	pDetInfo);
	void			OnKillFocusTemperature();
	BOOL			GetConfigFile(
						LPTSTR			szConfig,
						UINT			nBufferSize);
	long			GetNumberDetectors(
						CIniFile	*	pConfig);
	BOOL			GetDetectorName(
						CIniFile	*	pConfig,
						long			iDetector,
						LPTSTR			szName,
						UINT			nBufferSize);
	BOOL			GetDetectorPort(
						CIniFile	*	pConfig,
						long			iDetector,
						LPTSTR			szPort,
						UINT			nBufferSize);
	long			GetNumGains(
						CIniFile	*	pConfig,
						long			iDetector);
	BOOL			GetGainSetting(
						CIniFile	*	pConfig,
						long			iDetector,
						long			iGain,
						LPTSTR			szGain,
						UINT			nBufferSize);
	BOOL			GetHaveTempControl(
						CIniFile	*	pConfig,
						long			iDetector);
	long			GetNumChannels(
						CIniFile	*	pConfig,
						long			iDetector);
	long			GetChannelSetting(
						CIniFile	*	pConfig,
						long			iDetector,
						long			iChannel,
						double		*	lowLimit,
						double		*	hiLimit);
	BOOL			GetAppName(
						long			iDetector,
						LPTSTR			szApp,
						UINT			nBufferSize);
	void			InitDetectorInfo();
	void			ApplyDetector();
	BOOL			CheckReadOnly();			// check if the detector info is read only
	void			DisplayConfigChannelInfo(
						long			index);
	void			DisplayADInfoString();
	void			DisplayDetectorName();
	void			SetDetectorName(
						LPCTSTR			szName);
	void			InitGainSetting();
	void			DisplayGainSetting();
	void			SetGainSetting();
private:
	class CSingleDetector
	{
	public:
		CSingleDetector();
		~CSingleDetector();
		void			SetDetector(
							long			detector);
		BOOL			ReadConfig(
							CPropPageDetectorInfo * pPropPageDetectorInfo,
							CIniFile	*	pConfig);
		BOOL			CheckName(
							LPCTSTR			szName);
		void			GetName(
							LPTSTR			szName,
							UINT			nBufferSize);
		void			GetPort(
							LPTSTR			szPort,
							UINT			nBufferSize);
		long			GetNumGains();
		BOOL			GetGainString(
							long			iGain,
							LPTSTR			szGain,
							UINT			nBufferSize);
		long			FindGainIndex(
							LPCTSTR			szGain);
		BOOL			GethaveTempControl();
		long			GetNumChannels();
		long			GetADChannel(
							long			index);
		double			GetLowLimit(
							long			index);
		double			GetHiLimit(
							long			index);
		long			FindChannelIndex(
							long			channel);
	private:
		long			m_iDetector;
		TCHAR			m_szName[16];
		TCHAR			m_szPort[16];
		long			m_nGains;
		LPTSTR		*	m_arrGains;
		BOOL			m_haveTempControl;
		long			m_nChannels;
		long		*	m_arrChannels;
		double		*	m_arrLowLimits;
		double		*	m_arrHiLimits;
	};
	friend CSingleDetector;
	CSingleDetector*	m_pDetectors;
	long			m_nDetectors;
	long			m_currentDetector;
};
