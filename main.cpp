#include <MAPIX.h>
#include <MAPIUtil.h>
#include <string>
#include "AcctMgmt.h"
#include "AccountHelper.h"

#define	pbGlobalProfileSectionGuid	"\x13\xDB\xB0\xC8\xAA\x05\x10\x1A\x9B\xB0\x00\xAA\x00\x2F\xC4\x5A"

// This helper struct handles initializing and uninitializing COM for us
// Since it is in the global scope, it gets created before any of our main
// routine executes.
struct StartOle {
	StartOle() { CoInitialize(nullptr); }
	~StartOle() { CoUninitialize(); }
} _inst_StartOle;

std::wstring GetProfileName(LPMAPISESSION lpSession);
void IterateAllProps(LPOLKACCOUNT lpAccount);
HRESULT EnumerateAccounts(LPMAPISESSION lpSession, LPCWSTR lpwszProfile, bool bIterateAllProps);
HRESULT DisplayAccountList(LPMAPISESSION lpSession, LPCWSTR lpwszProfile, ULONG ulFlags);
void PrintBinary(DWORD cb, const BYTE* lpb);

// Execute a function, log and return the HRESULT
// Does not modify or reference existing hRes
#define EC_H(fnx) \
[&]() -> HRESULT { \
	auto __hRes = (fnx); \
	if (FAILED(__hRes)) printf("0x%08X from %s\n  @ %s + %d\n", __hRes, #fnx, __FILE__, __LINE__); \
	return __hRes; \
}()

void IterateAllProps(LPOLKACCOUNT lpAccount)
{
	if (!lpAccount) return;

	printf("Iterating all properties\r\n");
	auto hRes = S_OK;
	ULONG i = 0;
	ACCT_VARIANT pProp = { 0 };

	for (i = 0; i < 0x8000; i++)
	{
		memset(&pProp, 0, sizeof(ACCT_VARIANT));
		hRes = EC_H(lpAccount->GetProp(PROP_TAG(PT_LONG, i), &pProp));
		if (SUCCEEDED(hRes))
		{
			printf("Prop = 0x%08lX, Type = PT_LONG, Value = 0x%08lX\r\n", PROP_TAG(PT_LONG, i), pProp.Val.dw);
		}

		hRes = EC_H(lpAccount->GetProp(PROP_TAG(PT_UNICODE, i), &pProp));
		if (SUCCEEDED(hRes))
		{
			printf("Prop = 0x%08lX, Type = PT_UNICODE, Value = %ws\r\n", PROP_TAG(PT_UNICODE, i), pProp.Val.pwsz);
		}

		hRes = EC_H(lpAccount->GetProp(PROP_TAG(PT_BINARY, i), &pProp));
		if (SUCCEEDED(hRes))
		{
			printf("Prop = 0x%08lX, Type = PT_BINARY, Value = ", PROP_TAG(PT_BINARY, i));
			PrintBinary(pProp.Val.bin.cb, pProp.Val.bin.pb);
			printf("\r\n");
			(void)lpAccount->FreeMemory(static_cast<LPBYTE>(pProp.Val.bin.pb));
		}
	}

	printf("Done iterating all properties\r\n");
}

HRESULT EnumerateAccounts(LPMAPISESSION lpSession, LPCWSTR lpwszProfile, bool bIterateAllProps)
{
	auto hRes = S_OK;
	LPOLKACCOUNTMANAGER lpAcctMgr = nullptr;

	hRes = EC_H(CoCreateInstance(CLSID_OlkAccountManager,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IOlkAccountManager,
		reinterpret_cast<LPVOID*>(&lpAcctMgr)));
	if (SUCCEEDED(hRes) && lpAcctMgr)
	{
		auto pMyAcctHelper = new CAccountHelper(lpwszProfile, lpSession);
		if (pMyAcctHelper)
		{
			LPOLKACCOUNTHELPER lpAcctHelper = nullptr;
			hRes = EC_H(pMyAcctHelper->QueryInterface(IID_IOlkAccountHelper, reinterpret_cast<LPVOID*>(&lpAcctHelper)));
			if (SUCCEEDED(hRes) && lpAcctHelper)
			{
				hRes = EC_H(lpAcctMgr->Init(lpAcctHelper, ACCT_INIT_NOSYNCH_MAPI_ACCTS));
				if (SUCCEEDED(hRes))
				{
					LPOLKENUM lpAcctEnum = nullptr;

					hRes = EC_H(lpAcctMgr->EnumerateAccounts(&CLSID_OlkMail,
						nullptr,
						OLK_ACCOUNT_NO_FLAGS,
						&lpAcctEnum));
					if (SUCCEEDED(hRes) && lpAcctEnum)
					{
						DWORD cAccounts = 0;

						hRes = EC_H(lpAcctEnum->GetCount(&cAccounts));
						if (SUCCEEDED(hRes) && cAccounts)
						{
							hRes = EC_H(lpAcctEnum->Reset());
							if (SUCCEEDED(hRes))
							{
								DWORD i = 0;
								for (i = 0; i < cAccounts; i++)
								{
									if (i > 0) printf("\r\n");
									printf("Account #%lu\r\n", i + 1);
									LPUNKNOWN lpUnk = nullptr;

									hRes = EC_H(lpAcctEnum->GetNext(&lpUnk));
									if (SUCCEEDED(hRes) && lpUnk)
									{
										LPOLKACCOUNT lpAccount = nullptr;

										hRes = EC_H(lpUnk->QueryInterface(IID_IOlkAccount, reinterpret_cast<LPVOID*>(&lpAccount)));
										if (SUCCEEDED(hRes) && lpAccount)
										{
											ACCT_VARIANT pProp = { 0 };
											hRes = EC_H(lpAccount->GetProp(PROP_ACCT_NAME, &pProp));
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												printf("PROP_ACCT_NAME = \"%ws\"\r\n", pProp.Val.pwsz);
												(void)lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(pProp.Val.pwsz));
											}

											hRes = EC_H(lpAccount->GetProp(PROP_ACCT_USER_DISPLAY_NAME, &pProp));
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												printf("PROP_ACCT_USER_DISPLAY_NAME = \"%ws\"\r\n", pProp.Val.pwsz);
												(void)lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(pProp.Val.pwsz));
											}

											hRes = EC_H(lpAccount->GetProp(PROP_ACCT_USER_EMAIL_ADDR, &pProp));
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												printf("PROP_ACCT_USER_EMAIL_ADDR = \"%ws\"\r\n", pProp.Val.pwsz);
												(void)lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(pProp.Val.pwsz));
											}

											hRes = EC_H(lpAccount->GetProp(PROP_ACCT_STAMP, &pProp));
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												printf("PROP_ACCT_STAMP = \"%ws\"\r\n", pProp.Val.pwsz);
												(void)lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(pProp.Val.pwsz));
											}

											hRes = EC_H(lpAccount->GetProp(PROP_ACCT_SEND_STAMP, &pProp));
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												printf("PROP_ACCT_SEND_STAMP = \"%ws\"\r\n", pProp.Val.pwsz);
												(void)lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(pProp.Val.pwsz));
											}

											hRes = EC_H(lpAccount->GetProp(PROP_ACCT_IS_EXCH, &pProp));
											if (SUCCEEDED(hRes))
											{
												printf("PROP_ACCT_IS_EXCH = 0x%08lX\r\n", pProp.Val.dw);
											}

											hRes = EC_H(lpAccount->GetProp(PROP_ACCT_ID, &pProp));
											if (SUCCEEDED(hRes))
											{
												printf("PROP_ACCT_ID = 0x%08lX\r\n", pProp.Val.dw);
											}

											hRes = EC_H(lpAccount->GetProp(PROP_ACCT_DELIVERY_FOLDER, &pProp));
											if (SUCCEEDED(hRes) && pProp.Val.bin.cb && pProp.Val.bin.pb)
											{
												printf("PROP_ACCT_DELIVERY_FOLDER = ");
												PrintBinary(pProp.Val.bin.cb, pProp.Val.bin.pb);
												printf("\r\n");
												(void)lpAccount->FreeMemory(static_cast<LPBYTE>(pProp.Val.bin.pb));
											}

											hRes = EC_H(lpAccount->GetProp(PROP_ACCT_DELIVERY_STORE, &pProp));
											if (SUCCEEDED(hRes) && pProp.Val.bin.cb && pProp.Val.bin.pb)
											{
												printf("PROP_ACCT_DELIVERY_STORE = ");
												PrintBinary(pProp.Val.bin.cb, pProp.Val.bin.pb);
												printf("\r\n");
												(void)lpAccount->FreeMemory(static_cast<LPBYTE>(pProp.Val.bin.pb));
											}

											hRes = EC_H(lpAccount->GetProp(PROP_ACCT_SENTITEMS_EID, &pProp));
											if (SUCCEEDED(hRes) && pProp.Val.bin.cb && pProp.Val.bin.pb)
											{
												printf("PROP_ACCT_SENTITEMS_EID = ");
												PrintBinary(pProp.Val.bin.cb, pProp.Val.bin.pb);
												printf("\r\n");
												(void)lpAccount->FreeMemory(static_cast<LPBYTE>(pProp.Val.bin.pb));
											}

											hRes = EC_H(lpAccount->GetProp(PR_NEXT_SEND_ACCT, &pProp));
											if (SUCCEEDED(hRes) && pProp.Val.bin.cb && pProp.Val.bin.pb)
											{
												printf("PR_NEXT_SEND_ACCT = ");
												PrintBinary(pProp.Val.bin.cb, pProp.Val.bin.pb);
												printf("\r\n");
												(void)lpAccount->FreeMemory(static_cast<LPBYTE>(pProp.Val.bin.pb));
											}

											hRes = EC_H(lpAccount->GetProp(PR_PRIMARY_SEND_ACCT, &pProp));
											if (SUCCEEDED(hRes) && pProp.Val.bin.cb && pProp.Val.bin.pb)
											{
												printf("PR_PRIMARY_SEND_ACCT = ");
												PrintBinary(pProp.Val.bin.cb, pProp.Val.bin.pb);
												printf("\r\n");
												(void)lpAccount->FreeMemory(static_cast<LPBYTE>(pProp.Val.bin.pb));
											}

											hRes = EC_H(lpAccount->GetProp(PROP_MAPI_IDENTITY_ENTRYID, &pProp));
											if (SUCCEEDED(hRes) && pProp.Val.bin.cb && pProp.Val.bin.pb)
											{
												printf("PROP_MAPI_IDENTITY_ENTRYID = ");
												PrintBinary(pProp.Val.bin.cb, pProp.Val.bin.pb);
												printf("\r\n");
												(void)lpAccount->FreeMemory(static_cast<LPBYTE>(pProp.Val.bin.pb));
											}

											hRes = EC_H(lpAccount->GetProp(PROP_MAPI_TRANSPORT_FLAGS, &pProp));
											if (SUCCEEDED(hRes) && pProp.Val.bin.cb && pProp.Val.bin.pb)
											{
												printf("PROP_MAPI_TRANSPORT_FLAGS = ");
												PrintBinary(pProp.Val.bin.cb, pProp.Val.bin.pb);
												printf("\r\n");
												(void)lpAccount->FreeMemory(static_cast<LPBYTE>(pProp.Val.bin.pb));
											}
										}

										if (bIterateAllProps) IterateAllProps(lpAccount);
										if (lpAccount) lpAccount->Release();
										lpAccount = nullptr;
									}

									if (lpUnk) lpUnk->Release();
									lpUnk = nullptr;
								}
							}
						}
					}

					if (lpAcctEnum) lpAcctEnum->Release();
				}
			}

			if (lpAcctHelper) lpAcctHelper->Release();
		}

		if (pMyAcctHelper) pMyAcctHelper->Release();
	}

	if (lpAcctMgr) lpAcctMgr->Release();

	return hRes;
}

HRESULT DisplayAccountList(LPMAPISESSION lpSession, LPCWSTR lpwszProfile, ULONG ulFlags)
{
	auto hRes = S_OK;
	LPOLKACCOUNTMANAGER lpAcctMgr = nullptr;

	hRes = EC_H(CoCreateInstance(CLSID_OlkAccountManager,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IOlkAccountManager,
		reinterpret_cast<LPVOID*>(&lpAcctMgr)));
	if (SUCCEEDED(hRes) && lpAcctMgr)
	{
		auto pMyAcctHelper = new CAccountHelper(lpwszProfile, lpSession);
		if (pMyAcctHelper)
		{
			LPOLKACCOUNTHELPER lpAcctHelper = nullptr;
			hRes = EC_H(pMyAcctHelper->QueryInterface(IID_IOlkAccountHelper, reinterpret_cast<LPVOID*>(&lpAcctHelper)));
			if (SUCCEEDED(hRes) && lpAcctHelper)
			{
				hRes = EC_H(lpAcctMgr->Init(lpAcctHelper, ACCT_INIT_NOSYNCH_MAPI_ACCTS));
				if (SUCCEEDED(hRes))
				{
					hRes = EC_H(lpAcctMgr->DisplayAccountList(
						nullptr, // hwnd
						ulFlags, // dwFlags
						nullptr, // wszTitle
						NULL, // cCategories
						nullptr, // rgclsidCategories
						nullptr)); // pclsidType
				}
			}

			if (lpAcctHelper) lpAcctHelper->Release();
		}

		if (pMyAcctHelper) pMyAcctHelper->Release();
	}

	if (lpAcctMgr) lpAcctMgr->Release();

	return hRes;
}

std::wstring GetProfileName(LPMAPISESSION lpSession)
{
	auto hRes = S_OK;
	LPPROFSECT lpProfSect = nullptr;
	std::wstring profileName;

	if (!lpSession) return profileName;

	hRes = EC_H(lpSession->OpenProfileSection(
		LPMAPIUID(pbGlobalProfileSectionGuid),
		nullptr,
		0,
		&lpProfSect));
	if (SUCCEEDED(hRes) && lpProfSect)
	{
		LPSPropValue lpProfileName = nullptr;

		hRes = EC_H(HrGetOneProp(lpProfSect, PR_PROFILE_NAME_W, &lpProfileName));
		if (SUCCEEDED(hRes) && lpProfileName && lpProfileName->ulPropTag == PR_PROFILE_NAME_W)
		{
			profileName = std::wstring(lpProfileName->Value.lpszW);
		}

		MAPIFreeBuffer(lpProfileName);
	}

	if (lpProfSect)
		lpProfSect->Release();

	return profileName;
}

void PrintBinary(const DWORD cb, const BYTE* lpb)
{
	if (!cb || !lpb) return;
	LPSTR lpszHex = nullptr;
	ULONG i = 0;
	ULONG iBinPos = 0;
	lpszHex = new CHAR[1 + 2 * cb];
	if (lpszHex)
	{
		for (i = 0; i < cb; i++)
		{
			const auto bLow = static_cast<BYTE>(lpb[i] & 0xf);
			const auto bHigh = static_cast<BYTE>(lpb[i] >> 4 & 0xf);
			const auto szLow = static_cast<CHAR>(bLow <= 0x9 ? '0' + bLow : 'A' + bLow - 0xa);
			const auto szHigh = static_cast<CHAR>(bHigh <= 0x9 ? '0' + bHigh : 'A' + bHigh - 0xa);

			lpszHex[iBinPos] = szHigh;
			lpszHex[iBinPos + 1] = szLow;

			iBinPos += 2;
		}

		lpszHex[iBinPos] = '\0';
		printf("%hs", lpszHex);
		delete[] lpszHex;
	}
}

void DisplayUsage()
{
	printf("EnumAccounts - Exercise the account management API\n");
	printf("\n");
	printf("Usage:\n");
	printf("   EnumAccounts [-E] [-W] [-F]\n");
	printf("   EnumAccounts -?\n");
	printf("\n");
	printf("   -E   Enumerate accounts\n");
	printf("   -I   Iterate through all prop numbers looking for data (implies -E)\n");
	printf("   -W   Display wizard (DisplayAccountList)\n");
	printf("   -F   Indicates specific flags to pass to the DisplayAccountList.\n");
	printf("        If not specified, 0x500 = ACCTUI_NO_WARNING | ACCTUI_SHOW_ACCTWIZARD is used.\n");
	printf("        Available values are (these may be OR'ed together):\n");
	printf("           ACCTUI_NO_WARNING      0x0100\n");
	printf("           ACCTUI_SHOW_DATA_TAB   0x0200\n");
	printf("           ACCTUI_SHOW_ACCTWIZARD 0x0400\n");
	printf("   -?   Print this message\n");
	printf("\n");
	printf("Examples\n");
	printf("   EnumAccounts\n");
	printf("\n");
	printf("   EnumAccounts -e\n");
	printf("\n");
	printf("   EnumAccounts -w -f 0x500\n");
	printf("\n");
}

struct MYOPTIONS
{
	BOOL  bDoWizard;
	BOOL  bDoEnum;
	BOOL  bIterateAllProps;
	ULONG ulWizardFlags;
};

BOOL ParseArgs(int argc, char * argv[], MYOPTIONS * pRunOpts)
{
	if (!pRunOpts) return false;

	// Clear our options list
	ZeroMemory(pRunOpts, sizeof(MYOPTIONS));

	BOOL bFlagsSet = false;
	// Initialize non-null default values
	pRunOpts->ulWizardFlags = ACCTUI_NO_WARNING | ACCTUI_SHOW_ACCTWIZARD;

	for (auto i = 1; i < argc; i++)
	{
		switch (argv[i][0])
		{
		case '-':
		case '/':
		case '\\':
		{
			if (0 == argv[i][1])
			{
				// Bad argument - get out of here
				return false;
			}
			switch (tolower(argv[i][1]))
			{
			case 'i':
				pRunOpts->bIterateAllProps = true;
				pRunOpts->bDoEnum = true;
				break;
			case 'e':
				pRunOpts->bDoEnum = true;
				break;
			case 'w':
				pRunOpts->bDoWizard = true;
				break;
			case 'f':
				if (argc <= i + 1) return false;
				if (bFlagsSet) return false;
				{
					LPSTR szEndPtr = nullptr;
					pRunOpts->ulWizardFlags = strtoul(argv[++i], &szEndPtr, 16);
					bFlagsSet = true;
				}
				break;
			case '?':
			case 'h':
			default:
				// display help
				return FALSE;
			}
		}
		break;
		default:
			break;
		}
	}

	if (pRunOpts->bDoEnum && pRunOpts->bDoWizard) return false;
	if (pRunOpts->bIterateAllProps && pRunOpts->bDoWizard) return false;
	if (!pRunOpts->bDoEnum && !pRunOpts->bDoWizard) return false;

	// Didn't fail - return true
	return true;
}

void main(int argc, char * argv[])
{
	MYOPTIONS ProgOpts{};

	if (!ParseArgs(argc, argv, &ProgOpts))
	{
		DisplayUsage();
		return;
	}

	auto hRes = EC_H(MAPIInitialize(nullptr));
	if (SUCCEEDED(hRes))
	{
		LPMAPISESSION lpSession = nullptr;

		hRes = EC_H(MAPILogonEx(0,
			nullptr,
			nullptr,
			fMapiUnicode | MAPI_EXTENDED | MAPI_EXPLICIT_PROFILE |
			MAPI_NEW_SESSION | MAPI_NO_MAIL | MAPI_LOGON_UI,
			&lpSession));
		if (SUCCEEDED(hRes) && lpSession)
		{
			auto profileName = GetProfileName(lpSession);
			if (!profileName.empty())
			{
				if (ProgOpts.bDoWizard)
				{
					(void)DisplayAccountList(lpSession, profileName.c_str(), ProgOpts.ulWizardFlags);
				}
				else if (ProgOpts.bDoEnum)
				{
					(void)EnumerateAccounts(lpSession, profileName.c_str(), ProgOpts.bIterateAllProps);
				}
			}
		}

		if (lpSession) lpSession->Release();

		MAPIUninitialize();
	}
}