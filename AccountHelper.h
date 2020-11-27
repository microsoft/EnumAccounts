#pragma once

#include <MAPIX.h>
#include <string>
#include <AcctMgmt.h>

class CAccountHelper : IOlkAccountHelper
{
public:
	CAccountHelper(LPCWSTR lpwszProfName, LPMAPISESSION lpSession);
	~CAccountHelper();

	// IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, LPVOID * ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	// IOlkAccountHelper
	STDMETHODIMP PlaceHolder1(LPVOID)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP GetIdentity(LPWSTR pwszIdentity, DWORD * pcch);
	STDMETHODIMP GetMapiSession(LPUNKNOWN * ppmsess);
	STDMETHODIMP HandsOffSession();

private:
	LONG m_cRef;
	LPUNKNOWN m_lpUnkSession;
	std::wstring m_lpwszProfile;
};