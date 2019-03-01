
#include <iostream>
#include <iomanip>
#include <string>
#include <algorithm>
#include <fstream>
#include <atlstr.h>
#include <Shlobj.h>
#include <Windows.h>
#include <stdlib.h>
#include <dos.h>
#include <VersionHelpers.h>
#include <conio.h>
#include <direct.h>

using namespace std;

#ifndef UNICODE  
typedef std::string String;
#else
typedef std::wstring String;
#endif

//System Variables
TCHAR wWinDir[80]; String WinDir;
TCHAR wDataDir[80]; String DataDir;
TCHAR wPgmsDir[80]; String PgmsDir;

//ADOConnection Variables
TCHAR wConnectThru[80]; String ConnectThru;
TCHAR wDatabaseName[80]; String DatabaseName;
TCHAR wDataSource[80]; String DataSource;

//INSTALL Variables
TCHAR wSQLbuild[80]; String SQLbuild;
TCHAR wInstalledVersion[80]; String InstalledVersion;
TCHAR wBuild[80]; String Build;
TCHAR wServicePack[80]; String ServicePack;

String runningVersion = "v0.0.4";

int getOmate32();

void header();
void updateVisionWebINI();
void pause();

int main() {
	header();
	getOmate32();
	updateVisionWebINI();
	pause();
}

int getOmate32() {
	//cout << "Gathering Information from C:\\Windows\\Omate32.ini..." << string(2, '\n');

	//System Section
	GetPrivateProfileString(TEXT("System"), TEXT("WinDir"), TEXT("INI NOT FOUND"), wWinDir, 255, TEXT("Omate32.ini"));
	WinDir = wWinDir;
	GetPrivateProfileString(TEXT("System"), TEXT("DataDir"), TEXT("INI NOT FOUND"), wDataDir, 255, TEXT("Omate32.ini"));
	DataDir = wDataDir;
	GetPrivateProfileString(TEXT("System"), TEXT("PgmsDir"), TEXT("INI NOT FOUND"), wPgmsDir, 255, TEXT("Omate32.ini"));
	PgmsDir = wPgmsDir;

	//ADOConnection Section
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("ConnectThru"), TEXT("INI NOT FOUND"), wConnectThru, 255, TEXT("Omate32.ini"));
	ConnectThru = wConnectThru;
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("DatabaseName"), TEXT("INI NOT FOUND"), wDatabaseName, 255, TEXT("Omate32.ini"));
	DatabaseName = wDatabaseName;
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("DataSource"), TEXT("INI NOT FOUND"), wDataSource, 255, TEXT("Omate32.ini"));
	DataSource = wDataSource;

	//Install Section
	GetPrivateProfileString(TEXT("Install"), TEXT("SQL_Build"), TEXT("INI NOT FOUND"), wSQLbuild, 255, TEXT("Omate32.ini"));
	SQLbuild = wSQLbuild;
	GetPrivateProfileString(TEXT("Install"), TEXT("InstalledVersion"), TEXT("INI NOT FOUND"), wInstalledVersion, 255, TEXT("Omate32.ini"));
	InstalledVersion = wInstalledVersion;
	GetPrivateProfileString(TEXT("Install"), TEXT("Build"), TEXT("INI NOT FOUND"), wBuild, 255, TEXT("Omate32.ini"));
	Build = wBuild;
	GetPrivateProfileString(TEXT("Install"), TEXT("Service Pack"), TEXT("INI NOT FOUND"), wServicePack, 255, TEXT("Omate32.ini"));
	ServicePack = wServicePack;

	transform(WinDir.begin(), WinDir.end(), WinDir.begin(), (int(*)(int))toupper);
	transform(DataDir.begin(), DataDir.end(), DataDir.begin(), (int(*)(int))toupper);
	transform(PgmsDir.begin(), PgmsDir.end(), PgmsDir.begin(), (int(*)(int))toupper);

	transform(ConnectThru.begin(), ConnectThru.end(), ConnectThru.begin(), (int(*)(int))toupper);
	transform(DatabaseName.begin(), DatabaseName.end(), DatabaseName.begin(), (int(*)(int))toupper);
	transform(DataSource.begin(), DataSource.end(), DataSource.begin(), (int(*)(int))toupper);

	transform(SQLbuild.begin(), SQLbuild.end(), SQLbuild.begin(), (int(*)(int))toupper);
	transform(InstalledVersion.begin(), InstalledVersion.end(), InstalledVersion.begin(), (int(*)(int))toupper);
	transform(ServicePack.begin(), ServicePack.end(), ServicePack.begin(), (int(*)(int))toupper);

	if ((DataDir == "INI NOT FOUND") || (PgmsDir == "INI NOT FOUND")) {
		cout << setw(10) << left << "Omate32.ini not found in C:\\Windows. Please correct Omate32.ini to proceed" << endl;
		cout << setw(10) << left << "Is Officemate\\ExamWriter installed on this PC? Are you cloud hosted?";
		pause();
		exit(1);
	}

	return 0;
}

void header() {
	cout << left << setw(10) << "Eyefinity Officemate VisionWeb HTTPS Fix Tool " << runningVersion;
	cout << string(2, '\n');
}

void updateVisionWebINI() {

	// Possibly change the .reg file to be saved in a folder in the Officemate PgmsDir??? that way it could be run manually at a later time or if errored.
	string filenamepath = PgmsDir + "\\VisionWeb.ini";
	cout << left << setw(10) << "VisionWeb.ini Filepath = " << filenamepath << string(2, '\n');
	

	fstream VisionWebINI;
		VisionWebINI.open(filenamepath, fstream::out);
		VisionWebINI << "[VisionWeb]\n";
		VisionWebINI << "Available=Yes\n";
		VisionWebINI << "Integrated=Yes\n";
		VisionWebINI << "Graphic4Browser=VisionWeb.gif\n";
		VisionWebINI << "\n";

		VisionWebINI << "[VW Marketing Page]\n";
		VisionWebINI << "URL=https://www.visionweb.com/officemate/officemate.html\n";
		VisionWebINI << "\n";

		VisionWebINI << "[VWLogin]\n";
		VisionWebINI << "URL=https://www.visionweb.com/servlet/com.visionweb.login.servlets.LoginServlet\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Patient Spectacle Order]\n";
		VisionWebINI << "URL=https://www.visionweb.com/oerxlens/isapi/vwiisgeneric.dll/appentry.html\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Patient Contact Order]\n";
		VisionWebINI << "URL=https://www.visionweb.com/oerxsoft/isapi/vwiisgeneric.dll/appentry_rx.html\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Inventory Contact Order]\n";
		VisionWebINI << "URL=https://www.visionweb.com/oerxsoft/isapi/vwiisgeneric.dll/appentry_stk.html\n";
		VisionWebINI << "\n";

		VisionWebINI << ";THis url works for both patient and inventory(stock) orders\n";
		VisionWebINI << "[Patient Frame Order]\n";
		VisionWebINI << "URL=https://www.visionweb.com/frame/externalOrder.fr\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Inventory Frame Order]\n";
		VisionWebINI << "URL=https://www.visionweb.com/frame/externalOrder.fr\n";
		VisionWebINI << "\n";

		VisionWebINI << "[XML Headers]\n";
		VisionWebINI << "Hdr1_XMLVersion=<?xml version=\"1.0\"?>\n";
		VisionWebINI << "Hdr2_XMLStylesheetHref=<?xml:stylesheet type=\"text/xsl\" href=\"\n";
		VisionWebINI << ";The path to the xsl file goes between Hdr2 and Hdr25\n";
		VisionWebINI << "Header25_XMLHrefFName=VWDisplayOrder.xsl\"?>\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Order Status Request URL]\n";
		VisionWebINI << "URL_1=https://www.visionweb.com/oesc/isapi/vwiisgeneric.dll/api?GUID=\n";
		VisionWebINI << "URL_2=&ORDER=\n";
		VisionWebINI << "URL_3=&do=details&xml=0\n";
		VisionWebINI << ";;; 2 'GET' fields are needed: GUID(the sessionid) and ORDER(the order number)\n";
		VisionWebINI << ";;;\n";
		VisionWebINI << ";;; e.g.:\n";
		VisionWebINI << ";;; https://www.visionweb.com/oesc/isapi/vwiisgeneric.dll/api?GUID={73DEE315-1A\n";
		VisionWebINI << ";;; 4E-4905-B70D-6E30E57ACEBC}&ORDER=SP1b2z&do=details&xml=0\n";
		VisionWebINI << "\n";

		VisionWebINI << "[VisionWeb Time Out]\n";
		VisionWebINI << "Minutes=120\n";
		VisionWebINI << "\n";
		VisionWebINI << ";After an order is sent to pending or sent can show a msgbox with\n";
		VisionWebINI << ";  confirmation of status from return xml along with order number.\n";

		VisionWebINI << "[VisionWeb Status MsgBox]\n";
		VisionWebINI << "Show=No\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Status Pending Orders Page]\n";
		VisionWebINI << "URL=https://www.visionweb.com/oesc/isapi/vwiisgeneric.dll/ShoppingCart.html\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Status Sent Orders Page]\n";
		VisionWebINI << "URL=https://www.visionweb.com/oesc/isapi/vwiisgeneric.dll/Tracking.html\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Status Archived Orders Page]\n";
		VisionWebINI << "URL=https://www.visionweb.com/oesc/isapi/vwiisgeneric.dll/ArchiveSearch.html\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Order Number Page Text]\n";
		VisionWebINI << "URL_Text=Order\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Order Details Page Text]\n";
		VisionWebINI << "URL_Text=Details\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Order Sent Page Text]\n";
		VisionWebINI << "URL_Text=Send\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Order Delete Page Text]\n";
		VisionWebINI << "URL_Text=Delete\n";
		VisionWebINI << "\n";

		VisionWebINI << "[VW Error Msg]\n";
		VisionWebINI << "Message=\"Please contact VisionWeb Customer Service at CustomerService@VisionWeb.com or by calling 1-800-874-6601.\"";

		VisionWebINI.close();

	cout << left << setw(10) << "Finished Updating VisionWeb.ini..." << string(8, '\n');
	
}

void pause() {
	char a;
	cout << string(2, '\n') << "Press any Key to Continue... ";
	a = _getch();
}