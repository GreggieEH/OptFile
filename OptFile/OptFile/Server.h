#pragma once

class CServer
{
public:
	CServer(HINSTANCE hInst);
	~CServer(void);
	HINSTANCE			GetInstance();
	BOOL				CanUnloadNow();
	HRESULT				GetClassObject(
							REFCLSID	rclsid,
							REFIID		riid,
							LPVOID	*	ppv);
	HRESULT				RegisterServer();
	HRESULT				UnregisterServer();
	void				ObjectsUp();
	void				ObjectsDown();
	void				LocksUp();
	void				LocksDown();
	HRESULT				GetTypeLib(
							ITypeLib	**	ppTypeLib);
protected:
	BOOL				SetRegKeyValue(
							LPTSTR		pszKey,
							LPTSTR		pszSubkey,
							LPTSTR		pszValue);
	HRESULT				MyStringFromCLSID(
							REFGUID		refGuid,
							LPTSTR		szCLSID);

private:
	HINSTANCE			m_hInst;
	ULONG				m_cObjects;
	ULONG				m_cLocks;
};
