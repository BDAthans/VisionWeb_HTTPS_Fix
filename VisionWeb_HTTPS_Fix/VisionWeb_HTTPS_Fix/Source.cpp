
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

String runningVersion = "v0.0.7";

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
	cout << left << setw(10) << "Eyefinity Officemate VisionWeb HTTPS Login Fix Tool " << runningVersion;
	cout << string(2, '\n');
}

void updateVisionWebINI() {

	// Possibly change the .reg file to be saved in a folder in the Officemate PgmsDir??? that way it could be run manually at a later time or if errored.
	string filenamepath = PgmsDir + "\\VisionWeb.ini";
	cout << left << setw(10) << "VisionWeb.ini Filepath = " << filenamepath << string(2, '\n');
	

	fstream VisionWebINI;
		VisionWebINI.open(filenamepath, fstream::out);
		VisionWebINI << "[VisionWeb]";
		VisionWebINI << "\n";
		VisionWebINI << "Available=Yes";
		VisionWebINI << "\n";
		VisionWebINI << "Integrated=Yes";
		VisionWebINI << "\n";
		VisionWebINI << "Graphic4Browser=VisionWeb.gif";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[VW Marketing Page]";
		VisionWebINI << "\n";
		VisionWebINI << "URL=https://www.visionweb.com/officemate/officemate.html";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[VWLogin]";
		VisionWebINI << "\n";
		//VisionWebINI << "URL=https://www.visionweb.com/login/login.jsp";
		VisionWebINI << "URL=https://www.visionweb.com/servlet/com.visionweb.login.servlets.LoginServlet";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Patient Spectacle Order]";
		VisionWebINI << "\n";
		VisionWebINI << "URL=https://www.visionweb.com/oerxlens/isapi/vwiisgeneric.dll/appentry.html";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Patient Contact Order]";
		VisionWebINI << "\n";
		VisionWebINI << "URL=https://www.visionweb.com/oerxsoft/isapi/vwiisgeneric.dll/appentry_rx.html";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Inventory Contact Order]";
		VisionWebINI << "\n";
		VisionWebINI << "URL=https://www.visionweb.com/oerxsoft/isapi/vwiisgeneric.dll/appentry_stk.html";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << ";THis url works for both patient and inventory(stock) orders";
		VisionWebINI << "\n";
		VisionWebINI << "[Patient Frame Order]";
		VisionWebINI << "\n";
		VisionWebINI << "URL=https://www.visionweb.com/frame/externalOrder.fr";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Inventory Frame Order]";
		VisionWebINI << "\n";
		VisionWebINI << "URL=https://www.visionweb.com/frame/externalOrder.fr";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[XML Headers]";
		VisionWebINI << "\n";
		VisionWebINI << "Hdr1_XMLVersion=<?xml version=\"1.0\"?>";
		VisionWebINI << "\n";
		VisionWebINI << "Hdr2_XMLStylesheetHref=<?xml:stylesheet type=\"text/xsl\" href=\"";
		VisionWebINI << "\n";
		VisionWebINI << ";The path to the xsl file goes between Hdr2 and Hdr25";
		VisionWebINI << "\n";
		VisionWebINI << "Header25_XMLHrefFName=VWDisplayOrder.xsl\"?>";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Order Status Request URL]";
		VisionWebINI << "\n";
		VisionWebINI << "URL_1=https://www.visionweb.com/oesc/isapi/vwiisgeneric.dll/api?GUID=";
		VisionWebINI << "\n";
		VisionWebINI << "URL_2=&ORDER=";
		VisionWebINI << "\n";
		VisionWebINI << "URL_3=&do=details&xml=0";
		VisionWebINI << "\n";
		VisionWebINI << ";;; 2 'GET' fields are needed: GUID (the sessionid) and ORDER (the order number)";
		VisionWebINI << "\n";
		VisionWebINI << ";;;";
		VisionWebINI << "\n";
		VisionWebINI << ";;; e.g.:";
		VisionWebINI << "\n";
		VisionWebINI << ";;; https://www.visionweb.com/oesc/isapi/vwiisgeneric.dll/api?GUID={73DEE315-1A";
		VisionWebINI << "\n";
		VisionWebINI << ";;; 4E-4905-B70D-6E30E57ACEBC}&ORDER=SP1b2z&do=details&xml=0";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[VisionWeb Time Out]";
		VisionWebINI << "\n";
		VisionWebINI << "Minutes=120";
		VisionWebINI << "\n";
		VisionWebINI << "\n";
		VisionWebINI << ";After an order is sent to pending or sent can show a msgbox with";
		VisionWebINI << "\n";
		VisionWebINI << ";  confirmation of status from return xml along with order number.";
		VisionWebINI << "\n";

		VisionWebINI << "[VisionWeb Status MsgBox]";
		VisionWebINI << "\n";
		VisionWebINI << "Show=No";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Status Pending Orders Page]";
		VisionWebINI << "\n";
		VisionWebINI << "URL=https://www.visionweb.com/oesc/isapi/vwiisgeneric.dll/ShoppingCart.html";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Status Sent Orders Page]";
		VisionWebINI << "\n";
		VisionWebINI << "URL=https://www.visionweb.com/oesc/isapi/vwiisgeneric.dll/Tracking.html";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Status Archived Orders Page]";
		VisionWebINI << "\n";
		VisionWebINI << "URL=https://www.visionweb.com/oesc/isapi/vwiisgeneric.dll/ArchiveSearch.html";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Order Number Page Text]";
		VisionWebINI << "\n";
		VisionWebINI << "URL_Text=Order";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Order Details Page Text]";
		VisionWebINI << "\n";
		VisionWebINI << "URL_Text=Details";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Order Sent Page Text]";
		VisionWebINI << "\n";
		VisionWebINI << "URL_Text=Send";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[Order Delete Page Text]";
		VisionWebINI << "\n";
		VisionWebINI << "URL_Text=Delete";
		VisionWebINI << "\n";
		VisionWebINI << "\n";

		VisionWebINI << "[VW Error Msg]";
		VisionWebINI << "\n";
		VisionWebINI << "Message=\"Please contact VisionWeb Customer Service at CustomerService@VisionWeb.com or by calling 1-800-874-6601.\"";

		VisionWebINI.close();

	cout << left << setw(10) << "Finished Updating VisionWeb.ini..." << string(8, '\n');
	
}

void pause() {
	char a;
	cout << string(2, '\n') << "Press any Key to Continue... ";
	a = _getch();
}