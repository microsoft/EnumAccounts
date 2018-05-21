#pragma once

#include "AcctMgmt.h"
#include <MAPIX.h>
//#include <MAPIDefS.h>

class CAccountHelper : IOlkAccountHelper
{
public:
	CAccountHelper(LPWSTR lpwszProfName, LPMAPISESSION lpSession);
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
	LPWSTR m_lpwszProfile;
	size_t m_cchProfile;
};