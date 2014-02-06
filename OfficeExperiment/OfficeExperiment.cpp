// OfficeExperiment.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "OfficeExperiment.h"
#include <mmsystem.h>
#include <vector>
#include <iostream>

// link windows media lib
#pragma comment(lib, "winmm")

// Forward declarations of functions included in this code module:
void CycleTray(); 
BOOL RegisterMyProgramForStartup(LPCSTR pszAppName, LPCSTR pathToExe, LPCSTR args);

int APIENTRY _tWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPTSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
	TCHAR szPathToExe[MAX_PATH];
	GetModuleFileName(NULL, szPathToExe, MAX_PATH);
	RegisterMyProgramForStartup("Office_Experiement", szPathToExe, "");

	Sleep(3000);
	CycleTray();
}

BOOL RegisterMyProgramForStartup(LPCSTR pszAppName, LPCSTR pathToExe, LPCSTR args)
{
	HKEY hKey = NULL;
	LONG lResult = 0;
	BOOL fSuccess = TRUE;
	DWORD dwSize;

	const size_t count = MAX_PATH * 2;
	TCHAR szValue[count] = {};


	strcpy_s(szValue, count, "\"");
	strcat_s(szValue, count, pathToExe);
	strcat_s(szValue, count, "\" ");

	if (args != NULL)
	{
		// caller should make sure "args" is quoted if any single argument has a space
		// e.g. (L"-name \"Mark Voidale\"");
		strcat_s(szValue, count, args);
	}

	lResult = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0, NULL, 0, (KEY_WRITE | KEY_READ), NULL, &hKey, NULL);

	fSuccess = (lResult == 0);

	if (fSuccess)
	{
		dwSize = (strlen(szValue) + 1) * 2;
		lResult = RegSetValueEx(hKey, pszAppName, 0, REG_SZ, (BYTE*)szValue, dwSize);
		fSuccess = (lResult == 0);
	}

	if (hKey != NULL)
	{
		RegCloseKey(hKey);
		hKey = NULL;
	}

	return fSuccess;
}

static TCHAR driveLetters[26][5] =
{
	"A:\\",
	"B:\\",
	"C:\\",
	"D:\\",
	"E:\\",
	"F:\\",
	"G:\\",
	"H:\\",
	"I:\\",
	"J:\\",
	"K:\\",
	"L:\\",
	"M:\\",
	"N:\\",
	"O:\\",
	"P:\\",
	"Q:\\",
	"R:\\",
	"S:\\",
	"T:\\",
	"U:\\",
	"V:\\",
	"W:\\",
	"X:\\",
	"Y:\\",
	"Z:\\"
};

void ControlCdTray(TCHAR drive, DWORD command)
{
	// Not used here, only for debug
	MCIERROR mciError = 0;

	// Flags for MCI command
	DWORD mciFlags = MCI_WAIT | MCI_OPEN_SHAREABLE |
		MCI_OPEN_TYPE | MCI_OPEN_TYPE_ID | MCI_OPEN_ELEMENT;

	// Open drive device and get device ID
	TCHAR elementName[] = { drive };
	MCI_OPEN_PARMS mciOpenParms = { 0 };
	mciOpenParms.lpstrDeviceType = (LPCTSTR)MCI_DEVTYPE_CD_AUDIO;
	mciOpenParms.lpstrElementName = elementName;
	mciError = mciSendCommand(0,
		MCI_OPEN, mciFlags, (DWORD_PTR)&mciOpenParms);

	// Eject or close tray using device ID
	MCI_SET_PARMS mciSetParms = { 0 };
	mciFlags = MCI_WAIT | command; // command is sent by caller
	mciError = mciSendCommand(mciOpenParms.wDeviceID,
		MCI_SET, mciFlags, (DWORD_PTR)&mciSetParms);

	// Close device ID
	mciFlags = MCI_WAIT;
	MCI_GENERIC_PARMS mciGenericParms = { 0 };
	mciError = mciSendCommand(mciOpenParms.wDeviceID,
		MCI_CLOSE, mciFlags, (DWORD_PTR)&mciGenericParms);
}

// Eject drive tray
void EjectCdTray(TCHAR drive)
{
	ControlCdTray(drive, MCI_SET_DOOR_OPEN);
}

// Retract drive tray
void CloseCdTray(TCHAR drive)
{
	ControlCdTray(drive, MCI_SET_DOOR_CLOSED);
}

void CycleTray()
{
	DWORD drives = GetLogicalDrives();
	std::vector<std::string> cdDrives;

	for (DWORD i = 0; i < 26; ++i) //iterate through the alphabet
	{
		if (drives & (1 << i))
		{
			if (GetDriveType(driveLetters[i]) == DRIVE_CDROM)
				cdDrives.push_back(driveLetters[i]);
		}
	}

	std::string drivelist;
	for (std::string s : cdDrives)
	{
		drivelist.append(s + "\n");
	}

	//MessageBox(0, drivelist.c_str(), "CD Drives", MB_OK);
	
	bool eject = false;
	while (1)
	{
		eject = !eject;
		for (std::string s : cdDrives)
		{
			if (eject == true)
				EjectCdTray(s[0]);
			else
				CloseCdTray(s[0]);
		}
		Sleep(1000);
	}
}