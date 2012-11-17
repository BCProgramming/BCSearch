Attribute VB_Name = "MdlFileSystem"
Option Explicit
Public FileSystem As New BCFSObject
'Public Type FILETIME
'    dwLowDateTime As Long
'    dwHighDateTime As Long
'End Type
'Public Type PointAPI
'    x As Long
'    y As Long
'End Type
Private Declare Function DosDateTimeToFileTime Lib "kernel32.dll" (ByVal wFatDate As Long, ByVal wFatTime As Long, ByRef lpFileTime As FILETIME) As Long

Public mreg As New cRegistry
Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Public Declare Function PathCompactPathExA Lib "shlwapi.dll" (ByVal pszOut As String, ByVal pszSrc As String, ByVal cchmax As Long, ByVal dwFlags As Long) As Long
Public Declare Function PathCompactPathA Lib "shlwapi.dll" (ByVal hDC As Long, ByVal pszPath As String, ByVal dx As Long) As Long

Private Declare Function InternetConnectA Lib "wininet.dll" (ByRef hinternet As Long, ByVal lpszServerName As String, ByRef nServerPort As Long, ByVal lpszUserName As String, ByVal lpszPassword As String, ByVal dwService As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Long
Private Declare Function InternetConnectW Lib "wininet.dll" (ByRef hinternet As Long, ByVal lpszServerName As Long, ByRef nServerPort As Long, ByVal lpszUserName As Long, ByVal lpszPassword As Long, ByVal dwService As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Long

Private Declare Function InternetOpenA Lib "wininet.dll" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszproxy As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Private Declare Function InternetOpenW Lib "wininet.dll" (ByVal lpszAgent As Long, ByVal dwAccessType As Long, ByVal lpszproxy As Long, ByVal lpszProxyBypass As Long, ByVal dwFlags As Long) As Long

Private Declare Function GetCurrentDirectoryA Lib "kernel32.dll" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetCurrentDirectoryW Lib "kernel32.dll" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Private Declare Function SetCurrentDirectoryA Lib "kernel32.dll" (ByVal lpPathName As String) As Long
Private Declare Function SetCurrentDirectoryW Lib "kernel32.dll" (ByVal lpPathName As Long) As Long

Public Declare Function PathCompactPathExW Lib "shlwapi.dll" (ByVal pszOut As Long, ByVal pszSrc As Long, ByVal cchmax As Long, ByVal dwFlags As Long) As Long
Public Declare Function PathCompactPathW Lib "shlwapi.dll" (ByVal hDC As Long, ByVal pszPath As Long, ByVal dx As Long) As Long

Private Declare Function PathIsSameRootA Lib "shlwapi.dll" (ByVal pszPath1 As String, ByVal pszPath2 As String) As Long
Private Declare Function PathIsSameRootW Lib "shlwapi.dll" (ByVal pszPath1 As Long, ByVal pszPath2 As Long) As Long
'Known Folder APIS and constants...


'Constants first...
'// legacy CSIDL value: CSIDL_NETWORK
'// display name: "Network"
'// legacy display name: "My Network Places"
'// default path:
'// {D20BEEC4-5CA8-4905-AE3B-BF251EA09B53}
'DEFINE_KNOWN_FOLDER(FOLDERID_NetworkFolder, 0xD20BEEC4, 0x5CA8, 0x4905, 0xAE, 0x3B, 0xBF, 0x25, 0x1E, 0xA0, 0x9B, 0x53);
'
'// {0AC0837C-BBF8-452A-850D-79D08E667CA7}
'DEFINE_KNOWN_FOLDER(FOLDERID_ComputerFolder,   0x0AC0837C, 0xBBF8, 0x452A, 0x85, 0x0D, 0x79, 0xD0, 0x8E, 0x66, 0x7C, 0xA7);
'
'// {4D9F7874-4E0C-4904-967B-40B0D20C3E4B}
'DEFINE_KNOWN_FOLDER(FOLDERID_InternetFolder,      0x4D9F7874, 0x4E0C, 0x4904, 0x96, 0x7B, 0x40, 0xB0, 0xD2, 0x0C, 0x3E, 0x4B);
'
'// {82A74AEB-AEB4-465C-A014-D097EE346D63}
'DEFINE_KNOWN_FOLDER(FOLDERID_ControlPanelFolder,  0x82A74AEB, 0xAEB4, 0x465C, 0xA0, 0x14, 0xD0, 0x97, 0xEE, 0x34, 0x6D, 0x63);
'
'// {76FC4E2D-D6AD-4519-A663-37BD56068185}
'DEFINE_KNOWN_FOLDER(FOLDERID_PrintersFolder,      0x76FC4E2D, 0xD6AD, 0x4519, 0xA6, 0x63, 0x37, 0xBD, 0x56, 0x06, 0x81, 0x85);
'
'// {43668BF8-C14E-49B2-97C9-747784D784B7}
'DEFINE_KNOWN_FOLDER(FOLDERID_SyncManagerFolder,       0x43668BF8, 0xC14E, 0x49B2, 0x97, 0xC9, 0x74, 0x77, 0x84, 0xD7, 0x84, 0xB7);
'
'// {0F214138-B1D3-4a90-BBA9-27CBC0C5389A}
'DEFINE_KNOWN_FOLDER(FOLDERID_SyncSetupFolder, 0xf214138, 0xb1d3, 0x4a90, 0xbb, 0xa9, 0x27, 0xcb, 0xc0, 0xc5, 0x38, 0x9a);
'
'// {4bfefb45-347d-4006-a5be-ac0cb0567192}
'DEFINE_KNOWN_FOLDER(FOLDERID_ConflictFolder,      0x4bfefb45, 0x347d, 0x4006, 0xa5, 0xbe, 0xac, 0x0c, 0xb0, 0x56, 0x71, 0x92);
'
'// {289a9a43-be44-4057-a41b-587a76d7e7f9}
'DEFINE_KNOWN_FOLDER(FOLDERID_SyncResultsFolder,     0x289a9a43, 0xbe44, 0x4057, 0xa4, 0x1b, 0x58, 0x7a, 0x76, 0xd7, 0xe7, 0xf9);
'
'// {B7534046-3ECB-4C18-BE4E-64CD4CB7D6AC}
'DEFINE_KNOWN_FOLDER(FOLDERID_RecycleBinFolder,    0xB7534046, 0x3ECB, 0x4C18, 0xBE, 0x4E, 0x64, 0xCD, 0x4C, 0xB7, 0xD6, 0xAC);
'
'// {6F0CD92B-2E97-45D1-88FF-B0D186B8DEDD}
'DEFINE_KNOWN_FOLDER(FOLDERID_ConnectionsFolder,   0x6F0CD92B, 0x2E97, 0x45D1, 0x88, 0xFF, 0xB0, 0xD1, 0x86, 0xB8, 0xDE, 0xDD);
'
'// {FD228CB7-AE11-4AE3-864C-16F3910AB8FE}
'DEFINE_KNOWN_FOLDER(FOLDERID_Fonts,               0xFD228CB7, 0xAE11, 0x4AE3, 0x86, 0x4C, 0x16, 0xF3, 0x91, 0x0A, 0xB8, 0xFE);
'
'// display name:        "Desktop"
'// default path:        "C:\Users\<UserName>\Desktop"
'// legacy default path: "C:\Documents and Settings\<userName>\Desktop"
'// legacy CSIDL value:  CSIDL_DESKTOP
'// {B4BFCC3A-DB2C-424C-B029-7FE99A87C641}
'DEFINE_KNOWN_FOLDER(FOLDERID_Desktop,             0xB4BFCC3A, 0xDB2C, 0x424C, 0xB0, 0x29, 0x7F, 0xE9, 0x9A, 0x87, 0xC6, 0x41);
'
'// {B97D20BB-F46A-4C97-BA10-5E3608430854}
'DEFINE_KNOWN_FOLDER(FOLDERID_Startup,             0xB97D20BB, 0xF46A, 0x4C97, 0xBA, 0x10, 0x5E, 0x36, 0x08, 0x43, 0x08, 0x54);
'
'// {A77F5D77-2E2B-44C3-A6A2-ABA601054A51}
'DEFINE_KNOWN_FOLDER(FOLDERID_Programs,            0xA77F5D77, 0x2E2B, 0x44C3, 0xA6, 0xA2, 0xAB, 0xA6, 0x01, 0x05, 0x4A, 0x51);
'
'// {625B53C3-AB48-4EC1-BA1F-A1EF4146FC19}
'DEFINE_KNOWN_FOLDER(FOLDERID_StartMenu,           0x625B53C3, 0xAB48, 0x4EC1, 0xBA, 0x1F, 0xA1, 0xEF, 0x41, 0x46, 0xFC, 0x19);
'
'// {AE50C081-EBD2-438A-8655-8A092E34987A}
'DEFINE_KNOWN_FOLDER(FOLDERID_Recent,              0xAE50C081, 0xEBD2, 0x438A, 0x86, 0x55, 0x8A, 0x09, 0x2E, 0x34, 0x98, 0x7A);
'
'// {8983036C-27C0-404B-8F08-102D10DCFD74}
'DEFINE_KNOWN_FOLDER(FOLDERID_SendTo,              0x8983036C, 0x27C0, 0x404B, 0x8F, 0x08, 0x10, 0x2D, 0x10, 0xDC, 0xFD, 0x74);
'
'// {FDD39AD0-238F-46AF-ADB4-6C85480369C7}
'DEFINE_KNOWN_FOLDER(FOLDERID_Documents,           0xFDD39AD0, 0x238F, 0x46AF, 0xAD, 0xB4, 0x6C, 0x85, 0x48, 0x03, 0x69, 0xC7);
'
'// {1777F761-68AD-4D8A-87BD-30B759FA33DD}
'DEFINE_KNOWN_FOLDER(FOLDERID_Favorites,           0x1777F761, 0x68AD, 0x4D8A, 0x87, 0xBD, 0x30, 0xB7, 0x59, 0xFA, 0x33, 0xDD);
'
'// {C5ABBF53-E17F-4121-8900-86626FC2C973}
'DEFINE_KNOWN_FOLDER(FOLDERID_NetHood,             0xC5ABBF53, 0xE17F, 0x4121, 0x89, 0x00, 0x86, 0x62, 0x6F, 0xC2, 0xC9, 0x73);
'
'// {9274BD8D-CFD1-41C3-B35E-B13F55A758F4}
'DEFINE_KNOWN_FOLDER(FOLDERID_PrintHood,           0x9274BD8D, 0xCFD1, 0x41C3, 0xB3, 0x5E, 0xB1, 0x3F, 0x55, 0xA7, 0x58, 0xF4);
'
'// {A63293E8-664E-48DB-A079-DF759E0509F7}
'DEFINE_KNOWN_FOLDER(FOLDERID_Templates,           0xA63293E8, 0x664E, 0x48DB, 0xA0, 0x79, 0xDF, 0x75, 0x9E, 0x05, 0x09, 0xF7);
'
'// {82A5EA35-D9CD-47C5-9629-E15D2F714E6E}
'DEFINE_KNOWN_FOLDER(FOLDERID_CommonStartup,       0x82A5EA35, 0xD9CD, 0x47C5, 0x96, 0x29, 0xE1, 0x5D, 0x2F, 0x71, 0x4E, 0x6E);
'
'// {0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8}
'DEFINE_KNOWN_FOLDER(FOLDERID_CommonPrograms,      0x0139D44E, 0x6AFE, 0x49F2, 0x86, 0x90, 0x3D, 0xAF, 0xCA, 0xE6, 0xFF, 0xB8);
'
'// {A4115719-D62E-491D-AA7C-E74B8BE3B067}
'DEFINE_KNOWN_FOLDER(FOLDERID_CommonStartMenu,     0xA4115719, 0xD62E, 0x491D, 0xAA, 0x7C, 0xE7, 0x4B, 0x8B, 0xE3, 0xB0, 0x67);
'
'// {C4AA340D-F20F-4863-AFEF-F87EF2E6BA25}
'DEFINE_KNOWN_FOLDER(FOLDERID_PublicDesktop,       0xC4AA340D, 0xF20F, 0x4863, 0xAF, 0xEF, 0xF8, 0x7E, 0xF2, 0xE6, 0xBA, 0x25);
'
'// {62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}
'DEFINE_KNOWN_FOLDER(FOLDERID_ProgramData,         0x62AB5D82, 0xFDC1, 0x4DC3, 0xA9, 0xDD, 0x07, 0x0D, 0x1D, 0x49, 0x5D, 0x97);
'
'// {B94237E7-57AC-4347-9151-B08C6C32D1F7}
'DEFINE_KNOWN_FOLDER(FOLDERID_CommonTemplates,     0xB94237E7, 0x57AC, 0x4347, 0x91, 0x51, 0xB0, 0x8C, 0x6C, 0x32, 0xD1, 0xF7);
'
'// {ED4824AF-DCE4-45A8-81E2-FC7965083634}
'DEFINE_KNOWN_FOLDER(FOLDERID_PublicDocuments,     0xED4824AF, 0xDCE4, 0x45A8, 0x81, 0xE2, 0xFC, 0x79, 0x65, 0x08, 0x36, 0x34);
'
'// {3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}
'DEFINE_KNOWN_FOLDER(FOLDERID_RoamingAppData,      0x3EB685DB, 0x65F9, 0x4CF6, 0xA0, 0x3A, 0xE3, 0xEF, 0x65, 0x72, 0x9F, 0x3D);
'
'// {F1B32785-6FBA-4FCF-9D55-7B8E7F157091}
'DEFINE_KNOWN_FOLDER(FOLDERID_LocalAppData,        0xF1B32785, 0x6FBA, 0x4FCF, 0x9D, 0x55, 0x7B, 0x8E, 0x7F, 0x15, 0x70, 0x91);
'
'// {A520A1A4-1780-4FF6-BD18-167343C5AF16}
'DEFINE_KNOWN_FOLDER(FOLDERID_LocalAppDataLow,     0xA520A1A4, 0x1780, 0x4FF6, 0xBD, 0x18, 0x16, 0x73, 0x43, 0xC5, 0xAF, 0x16);
'
'// {352481E8-33BE-4251-BA85-6007CAEDCF9D}
'DEFINE_KNOWN_FOLDER(FOLDERID_InternetCache,       0x352481E8, 0x33BE, 0x4251, 0xBA, 0x85, 0x60, 0x07, 0xCA, 0xED, 0xCF, 0x9D);
'
'// {2B0F765D-C0E9-4171-908E-08A611B84FF6}
'DEFINE_KNOWN_FOLDER(FOLDERID_Cookies,             0x2B0F765D, 0xC0E9, 0x4171, 0x90, 0x8E, 0x08, 0xA6, 0x11, 0xB8, 0x4F, 0xF6);
'
'// {D9DC8A3B-B784-432E-A781-5A1130A75963}
'DEFINE_KNOWN_FOLDER(FOLDERID_History,             0xD9DC8A3B, 0xB784, 0x432E, 0xA7, 0x81, 0x5A, 0x11, 0x30, 0xA7, 0x59, 0x63);
'
'// {1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}
'DEFINE_KNOWN_FOLDER(FOLDERID_System,              0x1AC14E77, 0x02E7, 0x4E5D, 0xB7, 0x44, 0x2E, 0xB1, 0xAE, 0x51, 0x98, 0xB7);
'
'// {D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}
'DEFINE_KNOWN_FOLDER(FOLDERID_SystemX86,           0xD65231B0, 0xB2F1, 0x4857, 0xA4, 0xCE, 0xA8, 0xE7, 0xC6, 0xEA, 0x7D, 0x27);
'
'// {F38BF404-1D43-42F2-9305-67DE0B28FC23}
'DEFINE_KNOWN_FOLDER(FOLDERID_Windows,             0xF38BF404, 0x1D43, 0x42F2, 0x93, 0x05, 0x67, 0xDE, 0x0B, 0x28, 0xFC, 0x23);
'
'// {5E6C858F-0E22-4760-9AFE-EA3317B67173}
'DEFINE_KNOWN_FOLDER(FOLDERID_Profile,             0x5E6C858F, 0x0E22, 0x4760, 0x9A, 0xFE, 0xEA, 0x33, 0x17, 0xB6, 0x71, 0x73);
'
'// {33E28130-4E1E-4676-835A-98395C3BC3BB}
'DEFINE_KNOWN_FOLDER(FOLDERID_Pictures,            0x33E28130, 0x4E1E, 0x4676, 0x83, 0x5A, 0x98, 0x39, 0x5C, 0x3B, 0xC3, 0xBB);
'
'// {7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}
'DEFINE_KNOWN_FOLDER(FOLDERID_ProgramFilesX86,     0x7C5A40EF, 0xA0FB, 0x4BFC, 0x87, 0x4A, 0xC0, 0xF2, 0xE0, 0xB9, 0xFA, 0x8E);
'
'// {DE974D24-D9C6-4D3E-BF91-F4455120B917}
'DEFINE_KNOWN_FOLDER(FOLDERID_ProgramFilesCommonX86, 0xDE974D24, 0xD9C6, 0x4D3E, 0xBF, 0x91, 0xF4, 0x45, 0x51, 0x20, 0xB9, 0x17);
'
'// {6D809377-6AF0-444b-8957-A3773F02200E}
'DEFINE_KNOWN_FOLDER(FOLDERID_ProgramFilesX64,     0x6d809377, 0x6af0, 0x444b, 0x89, 0x57, 0xa3, 0x77, 0x3f, 0x02, 0x20, 0x0e );
'
'// {6365D5A7-0F0D-45e5-87F6-0DA56B6A4F7D}
'DEFINE_KNOWN_FOLDER(FOLDERID_ProgramFilesCommonX64, 0x6365d5a7, 0xf0d, 0x45e5, 0x87, 0xf6, 0xd, 0xa5, 0x6b, 0x6a, 0x4f, 0x7d );
'
'// {905e63b6-c1bf-494e-b29c-65b732d3d21a}
'DEFINE_KNOWN_FOLDER(FOLDERID_ProgramFiles,        0x905e63b6, 0xc1bf, 0x494e, 0xb2, 0x9c, 0x65, 0xb7, 0x32, 0xd3, 0xd2, 0x1a);
'
'// {F7F1ED05-9F6D-47A2-AAAE-29D317C6F066}
'DEFINE_KNOWN_FOLDER(FOLDERID_ProgramFilesCommon,  0xF7F1ED05, 0x9F6D, 0x47A2, 0xAA, 0xAE, 0x29, 0xD3, 0x17, 0xC6, 0xF0, 0x66);
'
'// {5cd7aee2-2219-4a67-b85d-6c9ce15660cb}
'DEFINE_KNOWN_FOLDER(FOLDERID_UserProgramFiles,    0x5cd7aee2, 0x2219, 0x4a67, 0xb8, 0x5d, 0x6c, 0x9c, 0xe1, 0x56, 0x60, 0xcb);
'
'// {bcbd3057-ca5c-4622-b42d-bc56db0ae516}
'DEFINE_KNOWN_FOLDER(FOLDERID_UserProgramFilesCommon, 0xbcbd3057, 0xca5c, 0x4622, 0xb4, 0x2d, 0xbc, 0x56, 0xdb, 0x0a, 0xe5, 0x16);
'
'// {724EF170-A42D-4FEF-9F26-B60E846FBA4F}
'DEFINE_KNOWN_FOLDER(FOLDERID_AdminTools,          0x724EF170, 0xA42D, 0x4FEF, 0x9F, 0x26, 0xB6, 0x0E, 0x84, 0x6F, 0xBA, 0x4F);
'
'// {D0384E7D-BAC3-4797-8F14-CBA229B392B5}
'DEFINE_KNOWN_FOLDER(FOLDERID_CommonAdminTools,    0xD0384E7D, 0xBAC3, 0x4797, 0x8F, 0x14, 0xCB, 0xA2, 0x29, 0xB3, 0x92, 0xB5);
'
'// {4BD8D571-6D19-48D3-BE97-422220080E43}
'DEFINE_KNOWN_FOLDER(FOLDERID_Music,               0x4BD8D571, 0x6D19, 0x48D3, 0xBE, 0x97, 0x42, 0x22, 0x20, 0x08, 0x0E, 0x43);
'
'// {18989B1D-99B5-455B-841C-AB7C74E4DDFC}
'DEFINE_KNOWN_FOLDER(FOLDERID_Videos,              0x18989B1D, 0x99B5, 0x455B, 0x84, 0x1C, 0xAB, 0x7C, 0x74, 0xE4, 0xDD, 0xFC);
'
'// {C870044B-F49E-4126-A9C3-B52A1FF411E8}
'DEFINE_KNOWN_FOLDER(FOLDERID_Ringtones,           0xC870044B, 0xF49E, 0x4126, 0xA9, 0xC3, 0xB5, 0x2A, 0x1F, 0xF4, 0x11, 0xE8);
'
'// {B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5}
'DEFINE_KNOWN_FOLDER(FOLDERID_PublicPictures,      0xB6EBFB86, 0x6907, 0x413C, 0x9A, 0xF7, 0x4F, 0xC2, 0xAB, 0xF0, 0x7C, 0xC5);
'
'// {3214FAB5-9757-4298-BB61-92A9DEAA44FF}
'DEFINE_KNOWN_FOLDER(FOLDERID_PublicMusic,         0x3214FAB5, 0x9757, 0x4298, 0xBB, 0x61, 0x92, 0xA9, 0xDE, 0xAA, 0x44, 0xFF);
'
'// {2400183A-6185-49FB-A2D8-4A392A602BA3}
'DEFINE_KNOWN_FOLDER(FOLDERID_PublicVideos,        0x2400183A, 0x6185, 0x49FB, 0xA2, 0xD8, 0x4A, 0x39, 0x2A, 0x60, 0x2B, 0xA3);
'
'// {E555AB60-153B-4D17-9F04-A5FE99FC15EC}
'DEFINE_KNOWN_FOLDER(FOLDERID_PublicRingtones,     0xE555AB60, 0x153B, 0x4D17, 0x9F, 0x04, 0xA5, 0xFE, 0x99, 0xFC, 0x15, 0xEC);
'
'// {8AD10C31-2ADB-4296-A8F7-E4701232C972}
'DEFINE_KNOWN_FOLDER(FOLDERID_ResourceDir,         0x8AD10C31, 0x2ADB, 0x4296, 0xA8, 0xF7, 0xE4, 0x70, 0x12, 0x32, 0xC9, 0x72);
'
'// {2A00375E-224C-49DE-B8D1-440DF7EF3DDC}
'DEFINE_KNOWN_FOLDER(FOLDERID_LocalizedResourcesDir, 0x2A00375E, 0x224C, 0x49DE, 0xB8, 0xD1, 0x44, 0x0D, 0xF7, 0xEF, 0x3D, 0xDC);
'
'// {C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D}
'DEFINE_KNOWN_FOLDER(FOLDERID_CommonOEMLinks,      0xC1BAE2D0, 0x10DF, 0x4334, 0xBE, 0xDD, 0x7A, 0xA2, 0x0B, 0x22, 0x7A, 0x9D);
'
'// {9E52AB10-F80D-49DF-ACB8-4330F5687855}
'DEFINE_KNOWN_FOLDER(FOLDERID_CDBurning,           0x9E52AB10, 0xF80D, 0x49DF, 0xAC, 0xB8, 0x43, 0x30, 0xF5, 0x68, 0x78, 0x55);
'
'// {0762D272-C50A-4BB0-A382-697DCD729B80}
'DEFINE_KNOWN_FOLDER(FOLDERID_UserProfiles,        0x0762D272, 0xC50A, 0x4BB0, 0xA3, 0x82, 0x69, 0x7D, 0xCD, 0x72, 0x9B, 0x80);
'
'// {DE92C1C7-837F-4F69-A3BB-86E631204A23}
'DEFINE_KNOWN_FOLDER(FOLDERID_Playlists,           0xDE92C1C7, 0x837F, 0x4F69, 0xA3, 0xBB, 0x86, 0xE6, 0x31, 0x20, 0x4A, 0x23);
'
'// {15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5}
'DEFINE_KNOWN_FOLDER(FOLDERID_SamplePlaylists,     0x15CA69B3, 0x30EE, 0x49C1, 0xAC, 0xE1, 0x6B, 0x5E, 0xC3, 0x72, 0xAF, 0xB5);
'
'// {B250C668-F57D-4EE1-A63C-290EE7D1AA1F}
'DEFINE_KNOWN_FOLDER(FOLDERID_SampleMusic,         0xB250C668, 0xF57D, 0x4EE1, 0xA6, 0x3C, 0x29, 0x0E, 0xE7, 0xD1, 0xAA, 0x1F);
'
'// {C4900540-2379-4C75-844B-64E6FAF8716B}
'DEFINE_KNOWN_FOLDER(FOLDERID_SamplePictures,      0xC4900540, 0x2379, 0x4C75, 0x84, 0x4B, 0x64, 0xE6, 0xFA, 0xF8, 0x71, 0x6B);
'
'// {859EAD94-2E85-48AD-A71A-0969CB56A6CD}
'DEFINE_KNOWN_FOLDER(FOLDERID_SampleVideos,        0x859EAD94, 0x2E85, 0x48AD, 0xA7, 0x1A, 0x09, 0x69, 0xCB, 0x56, 0xA6, 0xCD);
'
'// {69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C}
'DEFINE_KNOWN_FOLDER(FOLDERID_PhotoAlbums,         0x69D2CF90, 0xFC33, 0x4FB7, 0x9A, 0x0C, 0xEB, 0xB0, 0xF0, 0xFC, 0xB4, 0x3C);
'
'// {DFDF76A2-C82A-4D63-906A-5644AC457385}
'DEFINE_KNOWN_FOLDER(FOLDERID_Public,              0xDFDF76A2, 0xC82A, 0x4D63, 0x90, 0x6A, 0x56, 0x44, 0xAC, 0x45, 0x73, 0x85);
'
'// {df7266ac-9274-4867-8d55-3bd661de872d}
'DEFINE_KNOWN_FOLDER(FOLDERID_ChangeRemovePrograms,0xdf7266ac, 0x9274, 0x4867, 0x8d, 0x55, 0x3b, 0xd6, 0x61, 0xde, 0x87, 0x2d);
'
'// {a305ce99-f527-492b-8b1a-7e76fa98d6e4}
'DEFINE_KNOWN_FOLDER(FOLDERID_AppUpdates,          0xa305ce99, 0xf527, 0x492b, 0x8b, 0x1a, 0x7e, 0x76, 0xfa, 0x98, 0xd6, 0xe4);
'
'// {de61d971-5ebc-4f02-a3a9-6c82895e5c04}
'DEFINE_KNOWN_FOLDER(FOLDERID_AddNewPrograms,      0xde61d971, 0x5ebc, 0x4f02, 0xa3, 0xa9, 0x6c, 0x82, 0x89, 0x5e, 0x5c, 0x04);
'
'// {374DE290-123F-4565-9164-39C4925E467B}
'DEFINE_KNOWN_FOLDER(FOLDERID_Downloads,           0x374de290, 0x123f, 0x4565, 0x91, 0x64, 0x39, 0xc4, 0x92, 0x5e, 0x46, 0x7b);
'Const FOLDERID_Downloads = "{374DE290-123F-4565-9164-39C4925E467B}"
'// {3D644C9B-1FB8-4f30-9B45-F670235F79C0}
'DEFINE_KNOWN_FOLDER(FOLDERID_PublicDownloads,     0x3d644c9b, 0x1fb8, 0x4f30, 0x9b, 0x45, 0xf6, 0x70, 0x23, 0x5f, 0x79, 0xc0);
'
'// {7d1d3a04-debb-4115-95cf-2f29da2920da}
'DEFINE_KNOWN_FOLDER(FOLDERID_SavedSearches,       0x7d1d3a04, 0xdebb, 0x4115, 0x95, 0xcf, 0x2f, 0x29, 0xda, 0x29, 0x20, 0xda);
'
'// {52a4f021-7b75-48a9-9f6b-4b87a210bc8f}
'DEFINE_KNOWN_FOLDER(FOLDERID_QuickLaunch,         0x52a4f021, 0x7b75, 0x48a9, 0x9f, 0x6b, 0x4b, 0x87, 0xa2, 0x10, 0xbc, 0x8f);
'
'// {56784854-C6CB-462b-8169-88E350ACB882}
'DEFINE_KNOWN_FOLDER(FOLDERID_Contacts,            0x56784854, 0xc6cb, 0x462b, 0x81, 0x69, 0x88, 0xe3, 0x50, 0xac, 0xb8, 0x82);
'
'// {A75D362E-50FC-4fb7-AC2C-A8BEAA314493}
'DEFINE_GUID(FOLDERID_SidebarParts,                0xa75d362e, 0x50fc, 0x4fb7, 0xac, 0x2c, 0xa8, 0xbe, 0xaa, 0x31, 0x44, 0x93);
'
'// {7B396E54-9EC5-4300-BE0A-2482EBAE1A26}
'DEFINE_GUID(FOLDERID_SidebarDefaultParts,         0x7b396e54, 0x9ec5, 0x4300, 0xbe, 0xa, 0x24, 0x82, 0xeb, 0xae, 0x1a, 0x26);
'
'// {DEBF2536-E1A8-4c59-B6A2-414586476AEA}
'DEFINE_KNOWN_FOLDER(FOLDERID_PublicGameTasks,     0xdebf2536, 0xe1a8, 0x4c59, 0xb6, 0xa2, 0x41, 0x45, 0x86, 0x47, 0x6a, 0xea);
'
'// {054FAE61-4DD8-4787-80B6-090220C4B700}
'DEFINE_KNOWN_FOLDER(FOLDERID_GameTasks,           0x54fae61, 0x4dd8, 0x4787, 0x80, 0xb6, 0x9, 0x2, 0x20, 0xc4, 0xb7, 0x0);
'
'// {4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}
'DEFINE_KNOWN_FOLDER(FOLDERID_SavedGames,          0x4c5c32ff, 0xbb9d, 0x43b0, 0xb5, 0xb4, 0x2d, 0x72, 0xe5, 0x4e, 0xaa, 0xa4);
'
'// {CAC52C1A-B53D-4edc-92D7-6B2E8AC19434}
'DEFINE_KNOWN_FOLDER(FOLDERID_Games,               0xcac52c1a, 0xb53d, 0x4edc, 0x92, 0xd7, 0x6b, 0x2e, 0x8a, 0xc1, 0x94, 0x34);
'
'// {98ec0e18-2098-4d44-8644-66979315a281}
'DEFINE_KNOWN_FOLDER(FOLDERID_SEARCH_MAPI,         0x98ec0e18, 0x2098, 0x4d44, 0x86, 0x44, 0x66, 0x97, 0x93, 0x15, 0xa2, 0x81);
'
'// {ee32e446-31ca-4aba-814f-a5ebd2fd6d5e}
'DEFINE_KNOWN_FOLDER(FOLDERID_SEARCH_CSC,          0xee32e446, 0x31ca, 0x4aba, 0x81, 0x4f, 0xa5, 0xeb, 0xd2, 0xfd, 0x6d, 0x5e);
'
'// {bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968}
'DEFINE_KNOWN_FOLDER(FOLDERID_Links,               0xbfb9d5e0, 0xc6a9, 0x404c, 0xb2, 0xb2, 0xae, 0x6d, 0xb6, 0xaf, 0x49, 0x68);
'
'// {f3ce0f7c-4901-4acc-8648-d5d44b04ef8f}
'DEFINE_KNOWN_FOLDER(FOLDERID_UsersFiles,          0xf3ce0f7c, 0x4901, 0x4acc, 0x86, 0x48, 0xd5, 0xd4, 0x4b, 0x04, 0xef, 0x8f);
'
'// {A302545D-DEFF-464b-ABE8-61C8648D939B}
'DEFINE_KNOWN_FOLDER(FOLDERID_UsersLibraries,      0xa302545d, 0xdeff, 0x464b, 0xab, 0xe8, 0x61, 0xc8, 0x64, 0x8d, 0x93, 0x9b);
'
'// {190337d1-b8ca-4121-a639-6d472d16972a}
'DEFINE_KNOWN_FOLDER(FOLDERID_SearchHome,          0x190337d1, 0xb8ca, 0x4121, 0xa6, 0x39, 0x6d, 0x47, 0x2d, 0x16, 0x97, 0x2a);
'
'// {2C36C0AA-5812-4b87-BFD0-4CD0DFB19B39}
'DEFINE_KNOWN_FOLDER(FOLDERID_OriginalImages,      0x2C36C0AA, 0x5812, 0x4b87, 0xbf, 0xd0, 0x4c, 0xd0, 0xdf, 0xb1, 0x9b, 0x39);
'
'// {7b0db17d-9cd2-4a93-9733-46cc89022e7c}
'DEFINE_KNOWN_FOLDER(FOLDERID_DocumentsLibrary,    0x7b0db17d, 0x9cd2, 0x4a93, 0x97, 0x33, 0x46, 0xcc, 0x89, 0x02, 0x2e, 0x7c);
'
'// {2112AB0A-C86A-4ffe-A368-0DE96E47012E}
'DEFINE_KNOWN_FOLDER(FOLDERID_MusicLibrary,        0x2112ab0a, 0xc86a, 0x4ffe, 0xa3, 0x68, 0xd, 0xe9, 0x6e, 0x47, 0x1, 0x2e);
'
'// {A990AE9F-A03B-4e80-94BC-9912D7504104}
'DEFINE_KNOWN_FOLDER(FOLDERID_PicturesLibrary,     0xa990ae9f, 0xa03b, 0x4e80, 0x94, 0xbc, 0x99, 0x12, 0xd7, 0x50, 0x41, 0x4);
'
'// {491E922F-5643-4af4-A7EB-4E7A138D8174}
'DEFINE_KNOWN_FOLDER(FOLDERID_VideosLibrary,       0x491e922f, 0x5643, 0x4af4, 0xa7, 0xeb, 0x4e, 0x7a, 0x13, 0x8d, 0x81, 0x74);
'
'// {1A6FDBA2-F42D-4358-A798-B74D745926C5}
'DEFINE_KNOWN_FOLDER(FOLDERID_RecordedTVLibrary,   0x1a6fdba2, 0xf42d, 0x4358, 0xa7, 0x98, 0xb7, 0x4d, 0x74, 0x59, 0x26, 0xc5);
'
'// {52528A6B-B9E3-4add-B60D-588C2DBA842D}
'DEFINE_KNOWN_FOLDER(FOLDERID_HomeGroup,           0x52528a6b, 0xb9e3, 0x4add, 0xb6, 0xd, 0x58, 0x8c, 0x2d, 0xba, 0x84, 0x2d);
'
'// {5CE4A5E9-E4EB-479D-B89F-130C02886155}
'DEFINE_KNOWN_FOLDER(FOLDERID_DeviceMetadataStore, 0x5ce4a5e9, 0xe4eb, 0x479d, 0xb8, 0x9f, 0x13, 0x0c, 0x02, 0x88, 0x61, 0x55);
'
'// {1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}
'DEFINE_KNOWN_FOLDER(FOLDERID_Libraries,           0x1b3ea5dc, 0xb587, 0x4786, 0xb4, 0xef, 0xbd, 0x1d, 0xc3, 0x32, 0xae, 0xae);
'
'// {48daf80b-e6cf-4f4e-b800-0e69d84ee384}
'DEFINE_KNOWN_FOLDER(FOLDERID_PublicLibraries,     0x48daf80b, 0xe6cf, 0x4f4e, 0xb8, 0x00, 0x0e, 0x69, 0xd8, 0x4e, 0xe3, 0x84);
'
'// {9e3995ab-1f9c-4f13-b827-48b24b6c7174}
'DEFINE_KNOWN_FOLDER(FOLDERID_UserPinned,          0x9e3995ab, 0x1f9c, 0x4f13, 0xb8, 0x27, 0x48, 0xb2, 0x4b, 0x6c, 0x71, 0x74);
'
'// {bcb5256f-79f6-4cee-b725-dc34e402fd46}
'DEFINE_KNOWN_FOLDER(FOLDERID_ImplicitAppShortcuts,0xbcb5256f, 0x79f6, 0x4cee, 0xb7, 0x25, 0xdc, 0x34, 0xe4, 0x2, 0xfd, 0x46);



Public Declare Function SHGetKnownFolderPath Lib "shell32.dll" (rfid As olelib.UUID, ByVal dwFlags As Long, ByVal hToken As Long, ppszpath As Long) As Long

'HRESULT SHGetKnownFolderPath(
'  __in   REFKNOWNFOLDERID rfid,
'  __in   DWORD dwFlags,
'  __in   HANDLE hToken,
'  __out  PWSTR *ppszPath
');






Public Type BCCOPYFILEDATA
    BCCallback As IProgressCallback     'the callback.
    SourceFile As String        'used- the handles given in the callback don't have enough info....
    DestinationFile As String

End Type
Private Type WIN32_STREAM_ID
    dwStreamID As Long
    dwStreamAttributes As Long
    dwStreamSizeLow As Long
    dwStreamSizeHigh As Long
    dwStreamNameSize As Long
    'cStreamName As Byte
    'cStreamName() will IMMEDIATELY follow after reading this structure, then the stream data- which we should seek through, I suppose.
    
End Type
'FindStreamData.... For Windows Vista/Server 2003 Stream Enumeration functions.

'typedef struct _WIN32_FIND_STREAM_DATA {
'  LARGE_INTEGER StreamSize;
'  WCHAR         cStreamName[MAX_PATH + 36];
'}WIN32_FIND_STREAM_DATA, *PWIN32_FIND_STREAM_DATA;

Public Type ACL
    AclRevision As Byte
    Sbz1 As Byte
    AclSize As Integer
    AceCount As Integer
    Sbz2 As Integer
End Type


Public Type SECURITY_DESCRIPTOR
    Revision As Byte
    Sbz1 As Byte
    Control As Long
    Owner As Long
    Group As Long
    sAcl As ACL
    dacl As ACL
End Type


Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

'Vista/Server 2003 Stream Enumeration...
'HANDLE WINAPI FindFirstStreamW(
'  __in        LPCWSTR lpFileName,
'  __in        STREAM_INFO_LEVELS InfoLevel,
'  __out       LPVOID lpFindStreamData,
'  __reserved  DWORD dwFlags
');
'infolevel will be zero, for now. no other valid enumerations.
'Private Type LARGE_INTEGER
'    LoPart As Long
'    HiPart As Long
'End Type
Private Type WIN32_FIND_STREAM_DATA
    StreamSize As LARGE_INTEGER
    cStreamName As String * 296
End Type


Private Const ERROR_HANDLE_EOF As Long = 38&

Private Declare Function FindFirstStreamW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal Infolevel As Long, lpFindStreamData As WIN32_FIND_STREAM_DATA, ByVal dwFlags As Long) As Long
'BOOL WINAPI FindNextStreamW(
'  __in   HANDLE hFindStream,
'  __out  LPVOID lpFindStreamData
');
Private Declare Function FindNextStreamW Lib "kernel32.dll" (ByVal hFindStream As Long, lpFindStreamData As WIN32_FIND_STREAM_DATA) As Long

'HANDLE WINAPI FindFirstFileNameW(
'  __in     LPCWSTR lpFileName,
'  __in     DWORD dwFlags,
'  __inout  LPDWORD StringLength,
'  __inout  PWCHAR LinkName
');

Private Declare Function FindFirstFileNameW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwFlags As Long, ByVal StrLen As Long, ByVal LinkName As Long) As Long


'Public Declare Function GetFileAttributesEx Lib "kernel32.dll" Alias "GetFileAttributesExA" (ByVal lpFileName As String, ByVal fInfoLevelId As Struct_MembersOf_GET_FILEEX_INFO_LEVELS, ByRef lpFileInformation As Any) As Long
Private Declare Function CreateDirectoryA Lib "kernel32.dll" (ByVal lpPathName As String, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function RemoveDirectoryA Lib "kernel32.dll" (ByVal lpPathName As String) As Long
Private Declare Function PathCanonicalize Lib "shlwapi.dll" Alias "PathCanonicalizeA" (ByVal pszBuf As String, ByVal pszPath As String) As Long


Private Declare Function CreateDirectoryW Lib "kernel32.dll" (ByVal lpPathName As Long, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function RemoveDirectoryW Lib "kernel32.dll" (ByVal lpPathName As Long) As Long


'Public Declare Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, ByRef lpnLengthNeeded As Long) As Long
Public Declare Function GetFileType Lib "kernel32.dll" (ByVal hFile As Long) As Long

Public Declare Function GetFileAttributesA Lib "kernel32.dll" (ByVal lpFileName As String) As Long
Public Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long

Public Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSize As LARGE_INTEGER) As Long

Private Declare Function BackupRead Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, ByRef lpContext As Any) As Long
Private Declare Function BackupWrite Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, ByRef lpContext As Long) As Long
Private Declare Function BackupSeek Lib "kernel32.dll" (ByVal hFile As Long, ByVal dwLowBytesToSeek As Long, ByVal dwHighBytesToSeek As Long, ByRef lpdwLowByteSeeked As Long, ByRef lpdwHighByteSeeked As Long, ByRef lpContext As Any) As Long

Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Public Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long


' These WIDE/ANSI version are private. Their wrapper is made public.
Private Declare Function CreateFileA Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'
'Public Declare Function CreateDirectoryA Lib "kernel32" (ByVal lpPathName As String, lpSecurityAttributes As Any) As Long
'Public Declare Function CreateDirectoryW Lib "kernel32" (ByVal lpPathName As Long, lpSecurityAttributes As Any) As Long
'
'Private Declare Function GetFileAttributesA Lib "kernel32" (ByVal lpFileName As String) As Long
'Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long
'
'Private Declare Function SetFileAttributesA Lib "kernel32" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
'Private Declare Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
'
'Private Declare Function MoveFileA Lib "kernel32" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
'Private Declare Function MoveFileW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long
'
'Private Declare Function MoveFileExA Lib "kernel32" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
'Private Declare Function MoveFileExW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal dwFlags As Long) As Long
'
'Private Declare Function DeleteFileA Lib "kernel32" (ByVal lpFileName As String) As Long
'Private Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long
'
'Private Declare Function CreateDirectoryExA Lib "kernel32" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSECURITY_ATTRIBUTES As Any) As Long
'Private Declare Function CreateDirectoryExW Lib "kernel32" (ByVal lpTemplateDirectory As Long, ByVal lpNewDirectory As Long, lpSECURITY_ATTRIBUTES As Any) As Long
'
'Private Declare Function RemoveDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
'Private Declare Function RemoveDirectoryW Lib "kernel32" (ByVal lpPathName As Long) As Long
'
'Private Declare Function FindFirstFileA Lib "kernel32" (ByVal lpFileName As String, ByVal lpWIN32_FIND_DATA As Any) As Long
'Private Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal lpWIN32_FIND_DATA As Any) As Long
'
'Private Declare Function FindNextFileA Lib "kernel32" (ByVal hFindFile As Long, ByVal lpWIN32_FIND_DATA As Any) As Long
'Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, ByVal lpWIN32_FIND_DATA As Any) As Long
'
'Private Declare Function GetTempFileNameW Lib "kernel32" (ByVal lpszPath As Long, ByVal lpPrefixString As Long, ByVal wUnique As Long, ByVal lpTempFileName As Long) As Long
'Private Declare Function GetTempFileNameA Lib "kernel32" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
'
'Private Declare Function GetTempPathW Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
'Private Declare Function GetTempPathA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'
'Private Declare Function WNetGetConnectionA Lib "mpr.dll" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
'Private Declare Function WNetGetConnectionW Lib "mpr.dll" (ByVal lpszLocalName As Long, ByVal lpszRemoteName As Long, cbRemoteName As Long) As Long
'
'Private Declare Function GetVolumeInformationA Lib "kernel32" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'Private Declare Function GetVolumeInformationW Lib "kernel32" (ByVal lpRootPathName As Long, ByVal lpVolumeNameBuffer As Long, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As Long, ByVal nFileSystemNameSize As Long) As Long
'
'Private Declare Function GetLogicalDriveStringsA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Private Declare Function GetLogicalDriveStringsW Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long

Private Declare Function SHGetSpecialFolderPathA Lib "shell32.dll" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Private Declare Function SHGetSpecialFolderPathW Lib "shell32.dll" (ByVal hwnd As Long, ByVal pszPath As Long, ByVal csidl As Long, ByVal fCreate As Long) As Long

Private Declare Function SetFileAttributesA Lib "kernel32.dll" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function SetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Public Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Long

Private Declare Function CreateFileMappingA Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function CreateFileMappingALong Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As Long) As Long
Private Declare Function CreateFileMappingW Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileMappingAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpstrname As Long) As Long

Public Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByRef lpBaseAddress As Any) As Long



'The following require API wrappers:
Public Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetMappedFileName Lib "psapi.dll" Alias "GetMappedFileNameA" (ByVal hProcess As Long, ByRef lpv As Any, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function QueryDosDevice Lib "kernel32.dll" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long



Public Declare Function GetCursorPos Lib "user32" (Point As POINTAPI) As Long
Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long


'Public Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSize As LARGE_INTEGER) As Long
Private Const PROGRESS_CONTINUE As Long = 0
Private Const PROGRESS_CANCEL As Long = 1

Public Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type
Const MAX_PATH = 255

Public Type BCF_WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Type WIN32_FIND_DATAW
 
  dwFileAttributes        As Long
  ftCreationTime          As FILETIME
  ftLastAccessTime        As FILETIME
  ftLastWriteTime         As FILETIME

  nFileSizeHigh           As Long
  nFileSizeLow            As Long
  dwReserved0             As Long
  dwReserved1             As Long
  Buffer(1 To 240) As Byte
End Type
Public Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Public Declare Function FindFirstFileA Lib "kernel32.dll" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindFirstFileW Lib "kernel32.dll" (ByVal lpFileName As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFileA Lib "kernel32.dll" (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, ByRef lpWIN32_FIND_DATA As WIN32_FIND_DATA) As Long
'
Public Const ERROR_NO_MORE_FILES As Long = 18&
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" (ByVal hwnd As Long, ByVal csidl As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long
Private Declare Sub IIDFromString Lib "ole32.dll" (ByVal lpsz As String, ByVal lpiid As Long)

Public CDebug As New CDebug

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Sub CoCreateGuid Lib "ole32.dll" (ByRef pguid As Guid)
Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


Private Declare Function CloseHandleAPI Lib "kernel32.dll" Alias "CloseHandle" (ByVal hObject As Long) As Long
'Public Declare Function SetErrorMode Lib "kernel32.dll" (ByVal wMode As Long) As Long
'Public Declare Function CopyFileEx Lib "kernel32.dll" Alias "CopyFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByRef lpProgressRoutine As Long, ByRef lpData As Any, ByRef pbCancel As Long, ByVal dwCopyFlags As Long) As Long


Declare Function CopyFileEx Lib "kernel32.dll" Alias "CopyFileExW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal lpProgressRoutine As Long, lpData As Any, ByRef pbCancel As Long, ByVal dwCopyFlags As Long) As Long

'Public Declare Function CrePrivate Declare Function CopyFileEx Lib "kernel32.dll" Alias "CopyFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByRef lpProgressRoutine As PROGRESS_ROUTINE, ByRef lpData As Any, ByRef pbCancel As Long, ByVal dwCopyFlags As Long) As Long
Private Declare Function MoveFileWithProgress Lib "kernel32.dll" Alias "MoveFileWithProgressW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByRef lpProgressRoutine As Long, ByRef lpData As Any, ByVal dwFlags As Long) As Long


Public Declare Function GetFileInformationByHandle Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long

'Public Type SYSTEMTIME
'    wYear As Integer
'    wMonth As Integer
'    wDayOfWeek As Integer
'    wDay As Integer
'    wHour As Integer
'    wMinute As Integer
'    wSecond As Integer
'    wMilliseconds As Integer
'End Type

Public Declare Function SystemTimeToFileTime Lib "kernel32.dll" (ByRef lpSystemTime As SYSTEMTIME, ByRef lpFileTime As FILETIME) As Long
'Public Declare Function FileTimeToSystemTime Lib "kernel32.dll" (ByRef lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Long


Public Type OVERLAPPED
    internal As Long
    internalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type
Public Declare Function ReadFileEx Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpOverlapped As Long, ByVal lpCompletionRoutine As Long) As Long
Public Declare Function WriteFileEx Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long

Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Public Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)


Public Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function FlushFileBuffers Lib "kernel32.dll" (ByVal hFile As Long) As Long
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal length As Long)
Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000

'Public Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long


'Private Const MAX_PATH As Long = 260

Public Type BCF_SHFILEINFO
    hIcon As Long ' : icon
    iIcon As Long ' : icondex
    dwAttributes As Long ' : SFGAO_ flags
    szDisplayName As String * MAX_PATH ' : display name (or path)
    szTypeName As String * 80 ' : type name
End Type

Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Long
    hNameMaps As Long
    sProgress As String
End Type

Private Type ULARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

'Public Type SHFILEINFO
'    hIcon As Long ' : icon
'    iIcon As Long ' : icondex
'    dwAttributes As Long ' : SFGAO_ flags
'    szDisplayName As String * MAX_PATH ' : display name (or path)
'    szTypeName As String * 80 ' : type name
'End Type
'Public Type SHELLEXECUTEINFO
'    cbSize As Long
'    fMask As Long
'    hWnd As Long
'    lpVerb As String
'    lpFile As String
'    lpParameters As String
'    lpDirectory As String
'    nShow As Long
'    hInstApp As Long
'    ' fields
'    lpIDList As Long
'    lpClass As String
'    hkeyClass As Long
'    dwHotKey As Long
'    hIcon As Long
'    hProcess As Long
'End Type


'Public Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Long
Public Declare Sub SHEmptyRecycleBinA Lib "shell32.dll" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long)
Public Declare Sub SHEmptyRecycleBinW Lib "shell32.dll" (ByVal hwnd As Long, ByVal pszRootPath As Long, ByVal dwFlags As Long)



Public Declare Function CreateDirectoryExW Lib "kernel32.dll" (ByVal lpTemplateDirectory As Long, ByVal lpNewDirectory As Long, ByVal lpSecurityAttributes As Long) As Long
Public Declare Function CreateDirectoryExA Lib "kernel32.dll" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, ByVal lpSecurityAttributes As Long) As Long


Private Declare Function GetCompressedFileSizeA Lib "kernel32.dll" (ByVal lpFileName As String, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function GetCompressedFileSizeW Lib "kernel32.dll" (ByVal lpFileName As Long, ByRef lpFileSizeHigh As Long) As Long


'Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (ByRef lpFileOp As SHFILEOPSTRUCT) As Long
         Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long

Public Declare Function SHGetDiskFreeSpaceEx Lib "shell32.dll" Alias "SHGetDiskFreeSpaceExA" (ByVal pszDirectoryName As String, ByRef pulFreeBytesAvailableToCaller As ULARGE_INTEGER, ByRef pulTotalNumberOfBytes As ULARGE_INTEGER, ByRef pulTotalNumberOfFreeBytes As ULARGE_INTEGER) As Long
'Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'Error Code Value Meaning
Public Enum SHFileOperationErrors
    DE_SAMEFILE = &H71
'DE_SAMEFILE 0x71 The source and destination files are the same file.
    DE_MANYSRC1DEST = &H72
'DE_MANYSRC1DEST 0x72 Multiple file paths were specified in the source buffer, but only one destination file path.
    DE_DIFFDIR = &H73
'DE_DIFFDIR 0x73 Rename operation was specified but the destination path is a different directory. Use the move operation instead.
 DE_ROOTDIR = &H74
'DE_ROOTDIR 0x74 The source is a root directory, which cannot be moved or renamed.
 DE_OPCANCELLED = &H75
'DE_OPCANCELLED 0x75 The operation was cancelled by the user, or silently cancelled if the appropriate flags were supplied to SHFileOperation.
DE_DESTSUBTREE = &H76
'DE_DESTSUBTREE 0x76 The destination is a subtree of the source.
DE_ACCESSDENIEDSRC = &H78
'DE_ACCESSDENIEDSRC 0x78 Security settings denied access to the source.
DE_PATHTOODEEP = &H79
'DE_PATHTOODEEP 0x79 The source or destination path exceeded or would exceed MAX_PATH.
DE_MANYDEST = &H7A
'DE_MANYDEST 0x7A The operation involved multiple destination paths, which can fail in the case of a move operation.
DE_INVALIDFILES = &H7C
'DE_INVALIDFILES 0x7C The path in the source or destination or both was invalid.
DE_DESTSAMETREE = &H7D
'DE_DESTSAMETREE 0x7D The source and destination have the same parent folder.
DE_FLDDESTISFILE = &H7E
'DE_FLDDESTISFILE 0x7E The destination path is an existing file.
DE_FILEDESTISFLD = &H80
'DE_FILEDESTISFLD 0x80 The destination path is an existing folder.
DE_FILENAMETOOLONG = &H81
'DE_FILENAMETOOLONG 0x81 The name of the file exceeds MAX_PATH.
DE_DEST_IS_CDROM = &H82
'DE_DEST_IS_CDROM 0x82 The destination is a read-only CD-ROM, possibly unformatted.
DE_DEST_IS_DVD = &H83
'DE_DEST_IS_DVD 0x83 The destination is a read-only DVD, possibly unformatted.
DE_DEST_IS_CDRECORD = &H84
'DE_DEST_IS_CDRECORD 0x84 The destination is a writable CD-ROM, possibly unformatted.
DE_FILE_TOO_LARGE = &H85
'DE_FILE_TOO_LARGE 0x85 The file involved in the operation is too large for the destination media or file system.
DE_SRC_IS_CDROM = &H86
'DE_SRC_IS_CDROM 0x86 The source is a read-only CD-ROM, possibly unformatted.
DE_SRC_IS_DVD = &H87
'DE_SRC_IS_DVD 0x87 The source is a read-only DVD, possibly unformatted.
DE_SRC_IS_CDRECORD = &H88
'DE_SRC_IS_CDRECORD 0x88 The source is a writable CD-ROM, possibly unformatted.
DE_ERROR_MAX = &HB7

'DE_ERROR_MAX 0xB7 MAX_PATH was exceeded during the operation.

DE_UNKNOWN = &H402

ERRORONDEST = &H10000
' 0x402 An unknown error occurred. This is typically due to an invalid path in the source or destination. This error does not occur on Windows Vista and later.
'ERRORONDEST 0x10000 An unspecified error occurred on the destination.
'DE_ROOTDIR | ERRORONDEST 0x10074 Destination is a root directory and cannot be renamed.

End Enum

Public Const BCFileErrorBase = vbObjectError + 512 * 2 + 128 + 64
Public LargeIcons As cVBALImageList 'cache
Public ShellIcons As cVBALImageList 'cache
Public SmallIcons As cVBALImageList 'cache
Public Winmetrics As New SystemMetrics
'struct HANDLETOMAPPINGS
'{
'    UINT              uNumberOfMappings;  // Number of mappings in the array.
'    LPSHNAMEMAPPING   lpSHNameMapping;    // Pointer to the array of mappings.
'};
Public Declare Sub SHFreeNameMappings Lib "shell32.dll" (ByVal hNameMappings As Long)

Private Declare Sub SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwnd As Long, ByVal csidl As SpecialFolderConstants, ByVal ppidl As Long)

Public Type HANDLETOMAPPINGS
    uNumberOfMappings As Long
    LpFirstMapping As Long

End Type
Public Type SHNAMEMAPPING
    pszOldPath As String
    pszNewPath As String
    cchOldPath As Long
    cchNewPath As Long
End Type


Private Type MungeLong
    LongA As Long
    LongB As Long
End Type
Private Type MungeCurr
    CurrA As Currency
End Type
'
'Private ForceANSI As Boolean

Global ForceANSI As Boolean

Private Type IO_STATUS_BLOCK
IoStatus                As Long
Information             As Long
End Type

Private Const DATA_1 As String = "::$DATA"
Private Const DATA_2 As String = ":encryptable:$DATA"
Private Type FILE_STREAM_INFORMATION
    NextEntryOffset         As Long
    StreamNameLength        As Long
    StreamSize              As Long
    StreamSizeHi            As Long
    StreamAllocationSize    As Long
    StreamAllocationSizeHi  As Long
    StreamName(259)         As Byte
End Type

Private Const FileStreamInformation As Long = 22   ' from Enum FILE_INFORMATION_CLASS


'Public mdlFileSystem.TotalObjectCount As Long


'typedef struct tagOPENASINFO {
'    LPCWSTR pcszFile;
'    LPCWSTR pcszClass;
'    OPEN_AS_INFO_FLAGS oaifInFlags;
'} OPENASINFO;
'enum tagOPEN_AS_INFO_FLAGS {
'    OAIF_ALLOW_REGISTRATION = 0x00000001,     // enable the "always use this file" checkbox (NOTE if you don't pass this, it will be disabled)
'    OAIF_REGISTER_EXT       = 0x00000002,     // do the registration after the user hits "ok"
'    OAIF_EXEC               = 0x00000004,     // execute file after registering
'    OAIF_FORCE_REGISTRATION = 0x00000008,     // force the "always use this file" checkbox to be checked (normally, you won't use the OAIF_ALLOW_REGISTRATION when you pass this)
'#if (NTDDI_VERSION >= NTDDI_LONGHORN)
'    OAIF_HIDE_REGISTRATION  = 0x00000020,     // hide the "always use this file" checkbox
'    OAIF_URL_PROTOCOL       = 0x00000040,     // the "extension" passed is actually a protocol, and open with should show apps registered as capable of handling that protocol
'#End If
'};
'typedef int OPEN_AS_INFO_FLAGS;

Public Enum OPEN_AS_INFO_FLAGS
    OAIF_ALLOW_REGISTRATION = &H1
    OAIF_REGISTER_EXT = &H2
    OAIF_EXEC = &H4
    OAIF_FORCE_REGISTRATION = &H8
    'only in vista...
    OAIF_VISTA_HIDE_REGISTRATION = &H20
    OAIF_VISTA_URL_PROTOCOL = &H40
End Enum
Private Type OPENASINFO
    pcszFile As Long 'wide string.
    pcszClass As Long 'wide string.
    oaidInFlags As OPEN_AS_INFO_FLAGS
End Type
'
'HRESULT SHOpenWithDialog(
'    HWND hwndParent,
'    const OPENASINFO *poainfo
');
'shite- only supported in Vista...
Public Declare Function SHOpenWithDialog Lib "shell32" (ByVal hWndParent As Long, poaInfo As OPENASINFO) As Long

Public Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function NtQueryInformationFile Lib "NTDLL.DLL" (ByVal FileHandle As Long, IoStatusBlock_Out As IO_STATUS_BLOCK, lpFileInformation_Out As Long, ByVal length As Long, ByVal FileInformationClass As Long) As Long

Public Declare Function DeviceIoControl Lib "kernel32.dll" (ByVal hdevice As Long, ByVal dwIoControlCode As Long, ByRef lpInBuffer As Any, ByVal nInBufferSize As Long, ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, ByRef lpBytesReturned As Long, ByRef lpOverlapped As OVERLAPPED) As Long
'Public Declare Function DeviceIoControlAny Lib "kernel32.dll" Alias "DeviceIoControl" (ByVal hdevice As Long, ByVal dwIoControlCode As Long, ByRef lpInBuffer As Any, ByVal nInBufferSize As Long, ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, ByRef lpBytesReturned As Long, ByRef lpOverlapped As Any) As Long
    Public Declare Function DeviceIoControlAny Lib "kernel32" Alias "DeviceIoControl" (ByVal hdevice As Long, _
    ByVal dwIoControlCode As Long, lpInBuffer As Long, ByVal nInBufferSize As Integer, _
    lpOutBuffer As Long, ByVal nOutBufferSize As Long, lpBytesReturned As Long, _
    ByVal lpOverlapped As Any) As Long

Public Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Long

Private Declare Function QueryDosDeviceW Lib "kernel32.dll" (ByVal lpDeviceName As Long, ByVal lpTargetPath As Long, ByVal ucchMax As Long) As Long


Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwLen As Long, ByVal lpData As Long) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, ByRef lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (ByRef pBlock As Any, ByVal lpSubBlock As String, ByVal lplpBuffer As Long, ByRef puLen As Long) As Long



'volume functions

Private Declare Function FindFirstVolumeA Lib "kernel32.dll" (ByVal lpszVolumeName As String, ByVal cchBufferLength As Long) As Long
Private Declare Function FindFirstVolumeMountPointA Lib "kernel32.dll" (ByVal lpszRootPathName As String, ByVal lpszVolumeMountPoint As String, ByVal cchBufferLength As Long) As Long
Private Declare Function FindNextVolumeA Lib "kernel32.dll" (ByVal hFindVolume As Long, ByVal lpszVolumeName As String, ByVal cchBufferLength As Long) As Long
Private Declare Function FindNextVolumeMountPointA Lib "kernel32.dll" (ByVal hFindVolumeMountPoint As Long, ByVal lpszVolumeMountPoint As String, ByVal cchBufferLength As Long) As Long
Private Declare Function SetVolumeMountPointA Lib "kernel32.dll" (ByVal lpszVolumeMountPoint As String, ByVal lpszVolumeName As String) As Long
Private Declare Function GetVolumeNameForVolumeMountPointA Lib "kernel32.dll" (ByVal lpszVolumeMountPoint As String, ByVal lpszVolumeName As String, ByVal cchBufferLength As Long) As Long
Private Declare Function GetVolumePathNameA Lib "kernel32.dll" (ByVal lpszFileName As String, ByVal lpszVolumePathName As String, ByVal cchBufferLength As Long) As Long

Private Declare Function FindFirstVolumeW Lib "kernel32.dll" (ByVal lpszVolumeName As Long, ByVal cchBufferLength As Long) As Long
Private Declare Function FindFirstVolumeMountPointW Lib "kernel32.dll" (ByVal lpszRootPathName As Long, ByVal lpszVolumeMountPoint As Long, ByVal cchBufferLength As Long) As Long
Private Declare Function FindNextVolumeW Lib "kernel32.dll" (ByVal hFindVolume As Long, ByVal lpszVolumeName As Long, ByVal cchBufferLength As Long) As Long
Private Declare Function FindNextVolumeMountPointW Lib "kernel32.dll" (ByVal hFindVolumeMountPoint As Long, ByVal lpszVolumeMountPoint As Long, ByVal cchBufferLength As Long) As Long
Private Declare Function SetVolumeMountPointW Lib "kernel32.dll" (ByVal lpszVolumeMountPoint As Long, ByVal lpszVolumeName As Long) As Long
Private Declare Function GetVolumeNameForVolumeMountPointW Lib "kernel32.dll" (ByVal lpszVolumeMountPoint As Long, ByVal lpszVolumeName As Long, ByVal cchBufferLength As Long) As Long
Private Declare Function GetVolumePathNameW Lib "kernel32.dll" (ByVal lpszFileName As Long, ByVal lpszVolumePathName As Long, ByVal cchBufferLength As Long) As Long

'DWORD WINAPI GetFinalPathNameByHandle(
'  __in   HANDLE hFile,
'  __out  LPTSTR lpszFilePath,
'  __in   DWORD cchFilePath,
'  __in   DWORD dwFlags
');

Private Declare Function GetFinalPathNameByHandleA Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpszFilePath As String, ByVal cchFilePath As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetFinalPathNameByHandleW Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpszFilePath As Long, ByVal cchFilePath As Long, ByVal dwFlags As Long) As Long

Private Enum FileNamefromHandleFlags
    FILE_NAME_NORMALIZED = 0
    FILE_NAME_OPENED = &H8
    
    VOLUME_NAME_DOS = 0
    VOLUME_NAME_GUID = 1
    VOLUME_NAME_NONE = 4
    VOLUME_NAME_NT = 2
End Enum
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef lpExecInfo As SHELLEXECUTEINFOA) As Long
Private Type SHELLEXECUTEINFOA
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    ' fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long

Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByRef hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As String) As Long
Public Enum SetErrorModeConstants
 SEM_FAILCRITICALERRORS = &H1
 SEM_NOALIGNMENTFAULTEXCEPT = &H4
 SEM_NOGPFAULTERRORBOX = &H2
 SEM_NOOPENFILEERRORBOX = &H8000
End Enum
Private mWaitingAsync As Collection
'ReadFileEx Completion Routine:

Public Declare Function GetStdHandle Lib "kernel32.dll" (ByVal nStdHandle As Long) As Long
Public Const STD_INPUT_HANDLE As Long = -10&
Public Const STD_OUTPUT_HANDLE As Long = -11&
Public Const STD_ERROR_HANDLE As Long = -12&

'FileInformationEx...
'BOOL WINAPI GetFileInformationByHandleEx(
'  __in   HANDLE hFile,
'  __in   FILE_INFO_BY_HANDLE_CLASS FileInformationClass,
'  __out  LPVOID lpFileInformation,
'  __in   DWORD dwBufferSize
');

'Types/structures used by GFIBHEx:

'typedef struct _FILE_BASIC_INFO {
'  LARGE_INTEGER CreationTime;
'  LARGE_INTEGER LastAccessTime;
'  LARGE_INTEGER LastWriteTime;
'  LARGE_INTEGER ChangeTime;
'  DWORD         FileAttributes;
'} FILE_BASIC_INFO, *PFILE_BASIC_INFO;

'typedef struct _FILE_STANDARD_INFO {
'  LARGE_INTEGER AllocationSize;
'  LARGE_INTEGER EndOfFile;
'  DWORD          NumberOfLinks;
'  BOOL           DeletePending;
'  BOOL           Directory;
'} FILE_STANDARD_INFO, *PFILE_STANDARD_INFO;
'
'typedef struct _FILE_NAME_INFO {
'  DWORD FileNameLength;
'  WCHAR FileName[1];
'} FILE_NAME_INFO, *PFILE_NAME_INFO;

'typedef struct _FILE_STREAM_INFO {
'  DWORD         NextEntryOffset;
'  DWORD         StreamNameLength;
'  LARGE_INTEGER StreamSize;
'  LARGE_INTEGER StreamAllocationSize;
'  WCHAR         StreamName[1];
'} FILE_STREAM_INFO, *PFILE_STREAM_INFO;

'typedef struct _FILE_COMPRESSION_INFO {
'  LARGE_INTEGER CompressedFileSize;
'  WORD          CompressionFormat;
'  UCHAR         CompressionUnitShift;
'  UCHAR         ChunkShift;
'  UCHAR         ClusterShift;
'  UCHAR         Reserved[3];
'} FILE_COMPRESSION_INFO, *PFILE_COMPRESSION_INFO;

'typedef struct _FILE_ATTRIBUTE_TAG_INFO {
'  DWORD FileAttributes;
'  DWORD ReparseTag;
'} FILE_ATTRIBUTE_TAG_INFO, *PFILE_ATTRIBUTE_TAG_INFO;

'(for directory handles, used for e numeration, I think.
'typedef struct _FILE_ID_BOTH_DIR_INFO {
'  DWORD         NextEntryOffset;
'  DWORD         FileIndex;
'  LARGE_INTEGER CreationTime;
'  LARGE_INTEGER LastAccessTime;
'  LARGE_INTEGER LastWriteTime;
'  LARGE_INTEGER ChangeTime;
'  LARGE_INTEGER EndOfFile;
'  LARGE_INTEGER AllocationSize;
'  DWORD         FileAttributes;
'  DWORD         FileNameLength;
'  DWORD         EaSize;
'  CCHAR         ShortNameLength;
'  WCHAR         ShortName[12];
'  LARGE_INTEGER FileId;
'  WCHAR         FileName[1];
'} FILE_ID_BOTH_DIR_INFO, *PFILE_ID_BOTH_DIR_INFO;

'typedef struct _FILE_REMOTE_PROTOCOL_INFO {
'  USHORT StructureVersion;
'  USHORT StructureSize;
'  ULONG  Protocol;
'  USHORT ProtocolMajorVersion;
'  USHORT ProtocolMinorVersion;
'  USHORT ProtocolRevision;
'  USHORT Reserved;
'  ULONG  Flags;
'  struct {
'    ULONG Reserved[8];
'  } GenericReserved;
'  struct {
'    ULONG Reserved[16];
'  } ProtocolSpecificReserved;
'} FILE_REMOTE_PROTOCOL_INFO, *PFILE_REMOTE_PROTOCOL_INFO;
Public Enum KnownFolderIndexConstants
    FIDIDX_NetworkFolder = 0
    FIDIDX_ComputerFolder
    FIDIDX_InternetFolder
    FIDIDX_ControlPanelFolder
    FIDIDX_PrintersFolder
    FIDIDX_SyncManagerFolder
    FIDIDX_SyncSetupFolder
    FIDIDX_ConflictFolder
    FIDIDX_SyncResultsFolder
    FIDIDX_RecycleBinFolder
    FIDIDX_ConnectionsFolder
    FIDIDX_Fonts
    FIDIDX_Desktop
    FIDIDX_Startup
    FIDIDX_Programs
    FIDIDX_StartMenu
    FIDIDX_Recent
    FIDIDX_SendTo
    FIDIDX_Documents
    FIDIDX_Favorites
    FIDIDX_NetHood
    FIDIDX_PrintHood
    FIDIDX_Templates
    FIDIDX_CommonStartup
    FIDIDX_CommonPrograms
    FIDIDX_CommonStartMenu
    FIDIDX_PublicDesktop
    FIDIDX_ProgramData
    FIDIDX_CommonTemplates
    FIDIDX_PublicDocuments
    FIDIDX_RoamingAppData
    FIDIDX_LocalAppData
    FIDIDX_LocalAppDataLow
    FIDIDX_InternetCache
    FIDIDX_Cookies
    FIDIDX_History
    FIDIDX_System
    FIDIDX_SystemX86
    FIDIDX_Windows
    FIDIDX_Profile
    FIDIDX_Pictures
    FIDIDX_ProgramFilesX86
    FIDIDX_ProgramFilesCommonX86
    FIDIDX_ProgramFilesX64
    FIDIDX_ProgramFilesCommonX64
    FIDIDX_ProgramFiles
    FIDIDX_ProgramFilesCommon
    FIDIDX_AdminTools
    FIDIDX_CommonAdminTools
    FIDIDX_Music
    FIDIDX_Videos
    FIDIDX_PublicPictures
    FIDIDX_PublicMusic
    FIDIDX_PublicVideos
    FIDIDX_ResourceDir
    FIDIDX_LocalizedResourcesDir
    FIDIDX_CommonOEMLinks
    FIDIDX_CDBurning
    FIDIDX_UserProfiles
    FIDIDX_Playlists
    FIDIDX_SamplePlaylists
    FIDIDX_SampleMusic
    FIDIDX_SamplePictures
    FIDIDX_SampleVideos
    FIDIDX_PhotoAlbums
    FIDIDX_Public
    FIDIDX_ChangeRemovePrograms
    FIDIDX_AppUpdates
    FIDIDX_AddNewPrograms
    FIDIDX_Downloads
    FIDIDX_PublicDownloads
    FIDIDX_SavedSearches
    FIDIDX_QuickLaunch
    FIDIDX_Contacts
    FIDIDX_SidebarParts
    FIDIDX_SidebarDefaultParts
    FIDIDX_TreeProperties
    FIDIDX_PublicGameTasks
    FIDIDX_GameTasks
    FIDIDX_SavedGames
    FIDIDX_Games
    FIDIDX_RecordedTV
    FIDIDX_SEARCH_MAPI
    FIDIDX_SEARCH_CSC
    FIDIDX_Links
    FIDIDX_UsersFiles
    FIDIDX_SearchHome
    FIDIDX_OriginalImages
End Enum














Private Declare Function EncryptFileAPI Lib "advapi32.dll" Alias "EncryptFileW" (ByVal lpFileName As Long) As Long
Private Declare Function DecryptFileAPI Lib "advapi32.dll" Alias "DecryptFileW" (ByVal lpFileName As Long, ByVal dwReserved As Long) As Long
Dim knownfoldersGUIDs() As String
Dim KnownfoldersNames() As String

Dim mCopyProgresscallbacks() As IProgressCallback, mcopycallbackCount As Long
Private mTotalObjectCount  As Long
Dim createdstream
Private Declare Function GetFileInformationByHandleEx Lib "kernel32.dll" (ByVal hFile As Long, ByVal FileInformationClass As Long, ByVal lpFileInformation As Long, ByVal dwBufferSize As Long)



Private Declare Function GetFileSecurity Lib "advapi32.dll" Alias _
    "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation _
    As Long, pSecurityDescriptor As Byte, ByVal nLength As Long, _
    lpnLengthNeeded As Long) As Long
Private Declare Function GetSecurityDescriptorOwner Lib "advapi32.dll" _
    (pSecurityDescriptor As Any, pOwner As Long, lpbOwnerDefaulted As Long) As _
    Long
Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias _
    "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As Long, _
    ByVal Name As String, cbName As Long, ByVal ReferencedDomainName As String, _
    cbReferencedDomainName As Long, peUse As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias _
    "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long

Const OWNER_SECURITY_INFORMATION = &H1
Const ERROR_INSUFFICIENT_BUFFER = 122&


' return the name of the file owner
'
' runs over Windows NT or 2000, and works only with files in NTFS partitions

Function GetFileOwner(ByVal szfilename As String) As String
    Dim bSuccess As Long       ' Status variable
    Dim sizeSD As Long         ' Buffer size to store Owner's SID
    Dim pOwner As Long         ' Pointer to the Owner's SID
    Dim ownerName As String    ' Name of the file owner
    Dim domain_name As String  ' Name of the first domain for the owner
    Dim name_len As Long       ' Required length for the owner name
    Dim domain_len As Long     ' Required length for the domain name
    Dim sdBuf() As Byte        ' Buffer for Security Descriptor
    Dim nLength As Long        ' Length of the Windows Directory
    Dim deUse As Long          ' Pointer to a SID_NAME_USE enumerated type
                               ' indicating the type of the account
    
    ' Call GetFileSecurity the first time to obtain the size of the buffer
    ' required for the Security Descriptor.
    bSuccess = GetFileSecurity(szfilename, OWNER_SECURITY_INFORMATION, 0, 0&, _
        sizeSD)
    ' exit if any error
    If (bSuccess = 0) And (Err.LastDllError <> ERROR_INSUFFICIENT_BUFFER) Then _
        Exit Function
    
    ' Create a buffer of the required size and call GetFileSecurity again
    ReDim sdBuf(0 To sizeSD - 1) As Byte
    ' Fill the buffer with the security descriptor of the object specified by
    ' the
    ' filename parameter. The calling process must have the right to view the
    ' specified
    ' aspects of the object's security status.
    bSuccess = GetFileSecurity(szfilename, OWNER_SECURITY_INFORMATION, sdBuf(0), _
        sizeSD, sizeSD)
    ' exit if error
    If bSuccess = 0 Then Exit Function
    
    ' Obtain the owner's SID from the Security Descriptor, exit if error
    bSuccess = GetSecurityDescriptorOwner(sdBuf(0), pOwner, 0&)
    If bSuccess = 0 Then Exit Function

    ' Retrieve the name of the account and the name of the first domain on
    ' which this SID is found.  Passes in the Owner's SID obtained previously.
    ' Call LookupAccountSid twice, the
    ' first time to obtain the required size of the owner and domain names.
    bSuccess = LookupAccountSid(vbNullString, pOwner, ownerName, name_len, _
        domain_name, domain_len, deUse)
    ' exit if any error
    If (bSuccess = 0) And (Err.LastDllError <> ERROR_INSUFFICIENT_BUFFER) Then _
        Exit Function

    '  Allocate the required space in the name and domain_name string variables.
    ' Allocate 1 byte less to avoid the appended NULL character.
    ownerName = Space(name_len - 1)
    domain_name = Space(domain_len - 1)

    '  Call LookupAccountSid again to actually fill in the name of the owner
    ' and the first domain.
    bSuccess = LookupAccountSid(vbNullString, pOwner, ownerName, name_len, _
        domain_name, domain_len, deUse)
    If bSuccess = 0 Then Exit Function
       
    ' we've found a result
    GetFileOwner = ownerName
    
End Function





Public Property Let TotalObjectCount(ByVal Vdata As Long)
    mTotalObjectCount = Vdata
    Debug.Print "TotalObjectCount set to " & Vdata
End Property
Public Property Get TotalObjectCount() As Long
    TotalObjectCount = mTotalObjectCount
End Property
 
Public Function AddCPCallback(CallbackObject As IProgressCallback) As Long

    mcopycallbackCount = mcopycallbackCount + 1
    ReDim Preserve mCopyProgresscallbacks(mcopycallbackCount - 1)
    Set mCopyProgresscallbacks(mcopycallbackCount - 1) = CallbackObject
    AddCPCallback = mcopycallbackCount - 1

End Function
Public Sub RemoveCPCallback(ByVal CallbackCookie As Long)
    'removes a callback, given the cookie (which is an index into the array).
    
    Set mCopyProgresscallbacks(CallbackCookie) = Nothing
    
    
    'Now, the "fun" part
    
    'loop backwards through the array until we encounter a value that is not nothing:
    
    
    Dim I As Long, FirstSet As Long
    FirstSet = -1
    For I = UBound(mCopyProgresscallbacks) To 0 Step -1
    
        If Not mCopyProgresscallbacks(I) Is Nothing Then
            FirstSet = I
            Exit For
        End If
    Next
    If FirstSet = -1 Then
        'all of then are nothing. we can erase it completely.
        Erase mCopyProgresscallbacks
        mcopycallbackCount = 0
    Else
    'shrink the array to remove the unreferenced items...
        ReDim Preserve mCopyProgresscallbacks(FirstSet)
    
    End If


End Sub
Public Sub InitKnownFolders()
'
 '0x374de290, 0x123f, 0x4565, 0x91, 0x64, 0x39, 0xc4, 0x92, 0x5e, 0x46, 0x7b
 'knownfolders.FOLDERID_Downloads = DefGUID(&H374DE290, &H123F, &H4565, &H91, &H64, &H39, &HC4, &H92, &H5E, &H46, &H7B)
 
 Static flInitKnown As Boolean
 If Not flInitKnown Then
    flInitKnown = True
    ReDim knownfoldersGUIDs(88)
    ReDim KnownfoldersNames(88)
    knownfoldersGUIDs(0) = "{D20BEEC4-5CA8-4905-AE3B-BF251EA09B53}"
    knownfoldersGUIDs(1) = "{0AC0837C-BBF8-452A-850D-79D08E667CA7}"
    knownfoldersGUIDs(2) = "{4D9F7874-4E0C-4904-967B-40B0D20C3E4B}"
    knownfoldersGUIDs(3) = "{82A74AEB-AEB4-465C-A014-D097EE346D63}"
    knownfoldersGUIDs(4) = "{76FC4E2D-D6AD-4519-A663-37BD56068185}"
    knownfoldersGUIDs(5) = "{43668BF8-C14E-49B2-97C9-747784D784B7}"
    knownfoldersGUIDs(6) = "{0F214138-B1D3-4a90-BBA9-27CBC0C5389A}"
    knownfoldersGUIDs(7) = "{4bfefb45-347d-4006-a5be-ac0cb0567192}"
    knownfoldersGUIDs(8) = "{289a9a43-be44-4057-a41b-587a76d7e7f9}"
    knownfoldersGUIDs(9) = "{B7534046-3ECB-4C18-BE4E-64CD4CB7D6AC}"
    knownfoldersGUIDs(10) = "{6F0CD92B-2E97-45D1-88FF-B0D186B8DEDD}"
    knownfoldersGUIDs(11) = "{FD228CB7-AE11-4AE3-864C-16F3910AB8FE}"
    knownfoldersGUIDs(12) = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
    knownfoldersGUIDs(13) = "{B97D20BB-F46A-4C97-BA10-5E3608430854}"
    knownfoldersGUIDs(14) = "{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}"
    knownfoldersGUIDs(15) = "{625B53C3-AB48-4EC1-BA1F-A1EF4146FC19}"
    knownfoldersGUIDs(16) = "{AE50C081-EBD2-438A-8655-8A092E34987A}"
    knownfoldersGUIDs(17) = "{8983036C-27C0-404B-8F08-102D10DCFD74}"
    knownfoldersGUIDs(18) = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"
    knownfoldersGUIDs(19) = "{1777F761-68AD-4D8A-87BD-30B759FA33DD}"
    knownfoldersGUIDs(20) = "{C5ABBF53-E17F-4121-8900-86626FC2C973}"
    knownfoldersGUIDs(21) = "{9274BD8D-CFD1-41C3-B35E-B13F55A758F4}"
    knownfoldersGUIDs(22) = "{A63293E8-664E-48DB-A079-DF759E0509F7}"
    knownfoldersGUIDs(23) = "{82A5EA35-D9CD-47C5-9629-E15D2F714E6E}"
    knownfoldersGUIDs(24) = "{0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8}"
    knownfoldersGUIDs(25) = "{A4115719-D62E-491D-AA7C-E74B8BE3B067}"
    knownfoldersGUIDs(26) = "{C4AA340D-F20F-4863-AFEF-F87EF2E6BA25}"
    knownfoldersGUIDs(27) = "{62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}"
    knownfoldersGUIDs(28) = "{B94237E7-57AC-4347-9151-B08C6C32D1F7}"
    knownfoldersGUIDs(29) = "{ED4824AF-DCE4-45A8-81E2-FC7965083634}"
    knownfoldersGUIDs(30) = "{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}"
    knownfoldersGUIDs(31) = "{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}"
    knownfoldersGUIDs(32) = "{A520A1A4-1780-4FF6-BD18-167343C5AF16}"
    knownfoldersGUIDs(33) = "{352481E8-33BE-4251-BA85-6007CAEDCF9D}"
    knownfoldersGUIDs(34) = "{2B0F765D-C0E9-4171-908E-08A611B84FF6}"
    knownfoldersGUIDs(35) = "{D9DC8A3B-B784-432E-A781-5A1130A75963}"
    knownfoldersGUIDs(36) = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"
    knownfoldersGUIDs(37) = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}"
    knownfoldersGUIDs(38) = "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"
    knownfoldersGUIDs(39) = "{5E6C858F-0E22-4760-9AFE-EA3317B67173}"
    knownfoldersGUIDs(40) = "{33E28130-4E1E-4676-835A-98395C3BC3BB}"
    knownfoldersGUIDs(41) = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}"
    knownfoldersGUIDs(42) = "{DE974D24-D9C6-4D3E-BF91-F4455120B917}"
    knownfoldersGUIDs(43) = "{6D809377-6AF0-444b-8957-A3773F02200E}"
    knownfoldersGUIDs(44) = "{6365D5A7-0F0D-45e5-87F6-0DA56B6A4F7D}"
    knownfoldersGUIDs(45) = "{905e63b6-c1bf-494e-b29c-65b732d3d21a}"
    knownfoldersGUIDs(46) = "{F7F1ED05-9F6D-47A2-AAAE-29D317C6F066}"
    knownfoldersGUIDs(47) = "{724EF170-A42D-4FEF-9F26-B60E846FBA4F}"
    knownfoldersGUIDs(48) = "{D0384E7D-BAC3-4797-8F14-CBA229B392B5}"
    knownfoldersGUIDs(49) = "{4BD8D571-6D19-48D3-BE97-422220080E43}"
    knownfoldersGUIDs(50) = "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}"
    knownfoldersGUIDs(51) = "{B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5}"
    knownfoldersGUIDs(52) = "{3214FAB5-9757-4298-BB61-92A9DEAA44FF}"
    knownfoldersGUIDs(53) = "{2400183A-6185-49FB-A2D8-4A392A602BA3}"
    knownfoldersGUIDs(54) = "{8AD10C31-2ADB-4296-A8F7-E4701232C972}"
    knownfoldersGUIDs(55) = "{2A00375E-224C-49DE-B8D1-440DF7EF3DDC}"
    knownfoldersGUIDs(56) = "{C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D}"
    knownfoldersGUIDs(57) = "{9E52AB10-F80D-49DF-ACB8-4330F5687855}"
    knownfoldersGUIDs(58) = "{0762D272-C50A-4BB0-A382-697DCD729B80}"
    knownfoldersGUIDs(59) = "{DE92C1C7-837F-4F69-A3BB-86E631204A23}"
    knownfoldersGUIDs(60) = "{15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5}"
    knownfoldersGUIDs(61) = "{B250C668-F57D-4EE1-A63C-290EE7D1AA1F}"
    knownfoldersGUIDs(62) = "{C4900540-2379-4C75-844B-64E6FAF8716B}"
    knownfoldersGUIDs(63) = "{859EAD94-2E85-48AD-A71A-0969CB56A6CD}"
    knownfoldersGUIDs(64) = "{69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C}"
    knownfoldersGUIDs(65) = "{DFDF76A2-C82A-4D63-906A-5644AC457385}"
    knownfoldersGUIDs(66) = "{df7266ac-9274-4867-8d55-3bd661de872d}"
    knownfoldersGUIDs(67) = "{a305ce99-f527-492b-8b1a-7e76fa98d6e4}"
    knownfoldersGUIDs(68) = "{de61d971-5ebc-4f02-a3a9-6c82895e5c04}"
    knownfoldersGUIDs(69) = "{374DE290-123F-4565-9164-39C4925E467B}"
    knownfoldersGUIDs(70) = "{3D644C9B-1FB8-4f30-9B45-F670235F79C0}"
    knownfoldersGUIDs(71) = "{7d1d3a04-debb-4115-95cf-2f29da2920da}"
    knownfoldersGUIDs(72) = "{52a4f021-7b75-48a9-9f6b-4b87a210bc8f}"
    knownfoldersGUIDs(73) = "{56784854-C6CB-462b-8169-88E350ACB882}"
    knownfoldersGUIDs(74) = "{A75D362E-50FC-4fb7-AC2C-A8BEAA314493}"
    knownfoldersGUIDs(75) = "{7B396E54-9EC5-4300-BE0A-2482EBAE1A26}"
    knownfoldersGUIDs(76) = "{5b3749ad-b49f-49c1-83eb-15370fbd4882}"
    knownfoldersGUIDs(78) = "{DEBF2536-E1A8-4c59-B6A2-414586476AEA}"
    knownfoldersGUIDs(79) = "{054FAE61-4DD8-4787-80B6-090220C4B700}"
    knownfoldersGUIDs(80) = "{4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}"
    knownfoldersGUIDs(81) = "{CAC52C1A-B53D-4edc-92D7-6B2E8AC19434}"
    knownfoldersGUIDs(82) = "{bd85e001-112e-431e-983b-7b15ac09fff1}"
    knownfoldersGUIDs(83) = "{98ec0e18-2098-4d44-8644-66979315a281}"
    knownfoldersGUIDs(84) = "{ee32e446-31ca-4aba-814f-a5ebd2fd6d5e}"
    knownfoldersGUIDs(85) = "{bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968}"
    knownfoldersGUIDs(86) = "{f3ce0f7c-4901-4acc-8648-d5d44b04ef8f}"
    knownfoldersGUIDs(87) = "{190337d1-b8ca-4121-a639-6d472d16972a}"
    knownfoldersGUIDs(88) = "{2C36C0AA-5812-4b87-BFD0-4CD0DFB19B39}"
    
    
    
    'KnownfoldersNames(0)=
    
    
    KnownfoldersNames(0) = "NetworkFolder"
    KnownfoldersNames(1) = "ComputerFolder"
    KnownfoldersNames(2) = "InternetFolder"
    KnownfoldersNames(3) = "ControlPanelFolder"
    KnownfoldersNames(4) = "PrintersFolder"
    KnownfoldersNames(5) = "SyncManagerFolder"
    KnownfoldersNames(6) = "SyncSetupFolder"
    KnownfoldersNames(7) = "ConflictFolder"
    KnownfoldersNames(8) = "SyncResultsFolder"
    KnownfoldersNames(9) = "RecycleBinFolder"
    KnownfoldersNames(10) = "ConnectionsFolder"
    KnownfoldersNames(11) = "Fonts"
    KnownfoldersNames(12) = "Desktop"
    KnownfoldersNames(13) = "Startup"
    KnownfoldersNames(14) = "Programs"
    KnownfoldersNames(15) = "StartMenu"
    KnownfoldersNames(16) = "Recent"
    KnownfoldersNames(17) = "SendTo"
    KnownfoldersNames(18) = "Documents"
    KnownfoldersNames(19) = "Favorites"
    KnownfoldersNames(20) = "NetHood"
    KnownfoldersNames(21) = "PrintHood"
    KnownfoldersNames(22) = "Templates"
    KnownfoldersNames(23) = "CommonStartup"
    KnownfoldersNames(24) = "CommonPrograms"
    KnownfoldersNames(25) = "CommonStartMenu"
    KnownfoldersNames(26) = "PublicDesktop"
    KnownfoldersNames(27) = "ProgramData"
    KnownfoldersNames(28) = "CommonTemplates"
    KnownfoldersNames(29) = "PublicDocuments"
    KnownfoldersNames(30) = "RoamingAppData"
    KnownfoldersNames(31) = "LocalAppData"
    KnownfoldersNames(32) = "LocalAppDataLow"
    KnownfoldersNames(33) = "InternetCache"
    KnownfoldersNames(34) = "Cookies"
    KnownfoldersNames(35) = "History"
    KnownfoldersNames(36) = "System0"
    KnownfoldersNames(37) = "SystemX86"
    KnownfoldersNames(38) = "Windows"
    KnownfoldersNames(39) = "Profile"
    KnownfoldersNames(40) = "Pictures"
    KnownfoldersNames(41) = "ProgramFilesX86"
    KnownfoldersNames(42) = "ProgramFilesCommonX86"
    KnownfoldersNames(43) = "ProgramFilesX64"
    KnownfoldersNames(44) = "ProgramFilesCommonX64"
    KnownfoldersNames(45) = "ProgramFiles"
    KnownfoldersNames(46) = "ProgramFilesCommon"
    KnownfoldersNames(47) = "AdminTools"
    KnownfoldersNames(48) = "CommonAdminTools"
    KnownfoldersNames(49) = "Music"
    KnownfoldersNames(50) = "Videos"
    KnownfoldersNames(51) = "PublicPictures"
    KnownfoldersNames(52) = "PublicMusic"
    KnownfoldersNames(53) = "PublicVideos"
    KnownfoldersNames(54) = "ResourceDir"
    KnownfoldersNames(55) = "LocalizedResourcesDir"
    KnownfoldersNames(56) = "CommonOEMLinks"
    KnownfoldersNames(57) = "CDBurning"
    KnownfoldersNames(58) = "UserProfiles"
    KnownfoldersNames(59) = "Playlists"
    KnownfoldersNames(60) = "SamplePlaylists"
    KnownfoldersNames(61) = "SampleMusic"
    KnownfoldersNames(62) = "SamplePictures"
    KnownfoldersNames(63) = "SampleVideos"
    KnownfoldersNames(64) = "PhotoAlbums"
    KnownfoldersNames(65) = "Public"
    KnownfoldersNames(66) = "ChangeRemovePrograms"
    KnownfoldersNames(67) = "AppUpdates"
    KnownfoldersNames(68) = "AddNewPrograms"
    KnownfoldersNames(69) = "Downloads"
    KnownfoldersNames(70) = "PublicDownloads"
    KnownfoldersNames(71) = "SavedSearches"
    KnownfoldersNames(72) = "QuickLaunch"
    KnownfoldersNames(73) = "Contacts"
    KnownfoldersNames(74) = "SidebarParts"
    KnownfoldersNames(75) = "SidebarDefaultParts"
    KnownfoldersNames(76) = "TreeProperties"
    KnownfoldersNames(77) = "PublicGameTasks"
    KnownfoldersNames(78) = "GameTasks"
    KnownfoldersNames(79) = "SavedGames"
    KnownfoldersNames(80) = "Games"
    KnownfoldersNames(81) = "RecordedTV"
    KnownfoldersNames(82) = "SEARCH_MAPI"
    KnownfoldersNames(83) = "SEARCH_CSC"
    KnownfoldersNames(84) = "Links"
    KnownfoldersNames(85) = "UsersFiles"
    KnownfoldersNames(86) = "SearchHome"
    KnownfoldersNames(87) = "OriginalImages"
    KnownfoldersNames(88) = "Images"
    
    
    
    
    
End If
End Sub
Private Function GetKnownFolderGUID(EnumValue As KnownFolderIndexConstants) As String
    InitKnownFolders
    GetKnownFolderGUID = knownfoldersGUIDs(EnumValue)
End Function
Private Function GetKnownFolderStr(ByVal strCLSID As String) As String
    Dim retstr As String, retval As Long
    Dim ppszpath As Long
    Dim UUIDuse As olelib.UUID
    olelib.CLSIDFromString strCLSID, UUIDuse
        retval = SHGetKnownFolderPath(UUIDuse, 0, 0, ppszpath)
        If retval = 0 Then
      
       'if no error on return get the path
       'if present and release the pointer
         GetKnownFolderStr = GetPointerToByteStringW(ppszpath)
       ' GetKnownFolderStr = StringFromPointer(ppszpath)
         Call CoTaskMemFree(ppszpath)
         
      End If


    

End Function

Private Function GetPointerToByteStringW(ByVal dwData As Long) As String
  
   Dim tmp() As Byte
   Dim tmplen As Long
   
   If dwData <> 0 Then
   
     'determine the size of the returned data
      tmplen = lstrlenW(ByVal dwData) * 2
      
      If tmplen <> 0 Then
      
        'create a byte buffer for the string
        'then assign it to the return value
        'of the function
         ReDim tmp(0 To (tmplen - 1)) As Byte
         CopyMemory tmp(0), ByVal dwData, tmplen
         GetPointerToByteStringW = tmp
         
     End If
     
   End If
    
End Function
'Public Function DefGUID(ByVal longval As Long, ByVal word1 As Integer, ByVal word2 As Integer, ByVal b1 As Byte, ByVal b2 As Byte _
', ByVal b3 As Byte, ByVal b4 As Byte, ByVal b5 As Byte, ByVal b6 As Byte, ByVal b7 As Byte, ByVal b8 As Byte) As olelib.UUID
'Dim retuuid As olelib.UUID
'retuuid.Data1 = longval
'retuuid.Data2 = word1
'retuuid.Data3 = word2
'retuuid.Data4(0) = b1
'retuuid.Data4(1) = b2
'retuuid.Data4(2) = b3
'retuuid.Data4(3) = b4
'retuuid.Data4(4) = b5
'retuuid.Data4(5) = b6
'retuuid.Data4(6) = b7
'retuuid.Data4(7) = b8
'DefGUID = retuuid
''#define DEFINE_KNOWN_FOLDER(name, l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8)
'End Function

Public Function GetKnownFolder(ByVal KnownFolderEnum As KnownFolderIndexConstants)
    GetKnownFolder = GetKnownFolderStr(GetKnownFolderGUID(KnownFolderEnum))
End Function
Public Function GetKnownFolderDescription(ByVal KnownFolderEnum As KnownFolderIndexConstants) As String
    GetKnownFolderDescription = KnownfoldersNames(KnownFolderEnum)
End Function

'MapKnownFolderGUIDtoCSIDL
'used when a Pre-Vista OS is detected and the GetKnownFolder() Functions are called.
'not surprisingly, a large number of the Known folders will not have a corresponding CSIDL.

Public Function MapKnownFolderGUIDtoCSIDL(KnownFolderGUID As String) As CSIDLs
InitKnownFolders
Dim r As CSIDLs
'step one: cheat- convert the GUID into an index into the array, and then use the symbolic constants to determine the return value.
Dim I As Long, idxcompare As KnownFolderIndexConstants
For I = 0 To UBound(knownfoldersGUIDs)
    If StrComp(knownfoldersGUIDs(I), KnownFolderGUID, vbTextCompare) = 0 Then
        idxcompare = I
        Exit For
    
    End If

Next I
Select Case idxcompare
    Case FIDIDX_NetworkFolder = 0
        r = CSIDL_NETWORK
    Case FIDIDX_ComputerFolder
        r = CSIDL_DRIVES
    Case FIDIDX_InternetFolder
        r = CSIDL_INTERNET
    Case FIDIDX_ControlPanelFolder
        r = CSIDL_CONTROLS
    Case FIDIDX_PrintersFolder
        r = CSIDL_PRINTERS
    Case FIDIDX_SyncManagerFolder
        
    Case FIDIDX_SyncSetupFolder
    
    Case FIDIDX_ConflictFolder
    Case FIDIDX_SyncResultsFolder
    Case FIDIDX_RecycleBinFolder
        r = CSIDL_BITBUCKET
    Case FIDIDX_ConnectionsFolder
        r = CSIDL_CONNECTIONS
    Case FIDIDX_Fonts
        r = CSIDL_FONTS
    Case FIDIDX_Desktop
        r = CSIDL_DESKTOP
    Case FIDIDX_Startup
        r = CSIDL_STARTUP
    Case FIDIDX_Programs
        r = CSIDL_PROGRAM_FILES
    Case FIDIDX_StartMenu
        r = CSIDL_STARTMENU
    Case FIDIDX_Recent
        r = CSIDL_RECENT
    Case FIDIDX_SendTo
        r = CSIDL_SENDTO
    Case FIDIDX_Documents
        r = CSIDL_MYDOCUMENTS
    Case FIDIDX_Favorites
        r = CSIDL_FAVORITES
    Case FIDIDX_NetHood
        r = CSIDL_NETHOOD
    Case FIDIDX_PrintHood
        r = CSIDL_PRINTHOOD
    Case FIDIDX_Templates
        r = CSIDL_TEMPLATES
    Case FIDIDX_CommonStartup
        r = CSIDL_COMMON_STARTUP
    Case FIDIDX_CommonPrograms
        r = CSIDL_COMMON_PROGRAMS
    Case FIDIDX_CommonStartMenu
        r = CSIDL_COMMON_STARTMENU
    Case FIDIDX_PublicDesktop
        r = CSIDL_COMMON_DESKTOPDIRECTORY
    Case FIDIDX_ProgramData
        r = CSIDL_APPDATA
    Case FIDIDX_CommonTemplates
        r = CSIDL_COMMON_TEMPLATES
    Case FIDIDX_PublicDocuments
        r = CSIDL_COMMON_DOCUMENTS
    Case FIDIDX_RoamingAppData
        r = CSIDL_APPDATA
    Case FIDIDX_LocalAppData
        r = CSIDL_LOCAL_APPDATA
    Case FIDIDX_LocalAppDataLow
        r = CSIDL_COMMON_APPDATA
    Case FIDIDX_InternetCache
        r = CSIDL_INTERNET_CACHE
    Case FIDIDX_Cookies
        r = CSIDL_COOKIES
    Case FIDIDX_History
        r = CSIDL_HISTORY
    Case FIDIDX_System
        r = CSIDL_SYSTEM
    Case FIDIDX_SystemX86
        r = CSIDL_SYSTEMX86
    Case FIDIDX_Windows
        r = CSIDL_WINDOWS
    Case FIDIDX_Profile
        r = CSIDL_PROFILE
    Case FIDIDX_Pictures
        r = CSIDL_MYPICTURES
    Case FIDIDX_ProgramFilesX86
        r = CSIDL_PROGRAM_FILESX86
    Case FIDIDX_ProgramFilesCommonX86
        r = CSIDL_PROGRAM_FILES_COMMONX86
    Case FIDIDX_ProgramFilesX64
        r = CSIDL_PROGRAM_FILES
    Case FIDIDX_ProgramFilesCommonX64
        r = CSIDL_PROGRAM_FILES_COMMON
    Case FIDIDX_ProgramFiles
        r = CSIDL_PROGRAM_FILES
    Case FIDIDX_ProgramFilesCommon
        r = CSIDL_PROGRAM_FILES_COMMON
    Case FIDIDX_AdminTools
        r = CSIDL_ADMINTOOLS
    Case FIDIDX_CommonAdminTools
        r = CSIDL_COMMON_ADMINTOOLS
    Case FIDIDX_Music
        r = CSIDL_MYMUSIC
    Case FIDIDX_Videos
        r = CSIDL_MYVIDEO
    Case FIDIDX_PublicPictures
        r = CSIDL_COMMON_PICTURES
    Case FIDIDX_PublicMusic
        r = CSIDL_COMMON_MUSIC
    Case FIDIDX_PublicVideos
        r = CSIDL_COMMON_VIDEO
    Case FIDIDX_ResourceDir
        r = CSIDL_RESOURCES
    Case FIDIDX_LocalizedResourcesDir
        r = CSIDL_RESOURCES_LOCALIZED
    Case FIDIDX_CommonOEMLinks
        r = CSIDL_COMMON_OEM_LINKS
    Case FIDIDX_CDBurning
        r = CSIDL_CDBURN_AREA
    Case FIDIDX_UserProfiles
        r = CSIDL_PROFILE
    Case FIDIDX_Playlists
        r = CSIDL_MYMUSIC
    Case FIDIDX_SamplePlaylists
    Case FIDIDX_SampleMusic
    Case FIDIDX_SamplePictures
    Case FIDIDX_SampleVideos
    Case FIDIDX_PhotoAlbums
    Case FIDIDX_Public
        r = CSIDL_COMMON_DOCUMENTS
    Case FIDIDX_ChangeRemovePrograms
    
    Case FIDIDX_AppUpdates
    Case FIDIDX_AddNewPrograms
    Case FIDIDX_Downloads
        
    Case FIDIDX_PublicDownloads
    Case FIDIDX_SavedSearches
    Case FIDIDX_QuickLaunch
    Case FIDIDX_Contacts
    Case FIDIDX_SidebarParts
    Case FIDIDX_SidebarDefaultParts
    Case FIDIDX_TreeProperties
    Case FIDIDX_PublicGameTasks
    Case FIDIDX_GameTasks
    Case FIDIDX_SavedGames
    Case FIDIDX_Games
    Case FIDIDX_RecordedTV
    Case FIDIDX_SEARCH_MAPI
    Case FIDIDX_SEARCH_CSC
    Case FIDIDX_Links
    Case FIDIDX_UsersFiles
        r = CSIDL_PROFILE
    Case FIDIDX_SearchHome
    Case FIDIDX_OriginalImages


End Select

End Function
Public Sub TestKnownFolders()
'FOLDERID_Downloads
Dim strpath As String
strpath = GetKnownFolderStr(GetKnownFolderGUID(FIDIDX_OriginalImages))

End Sub
Public Function EncryptFile(ByVal lpFileName As String) As Long
    EncryptFile = EncryptFileAPI(StrPtr(lpFileName))
End Function
Public Function DecryptFile(ByVal lpFileName As String, dwReserved As Long)
    DecryptFile = DecryptFileAPI(StrPtr(lpFileName), dwReserved)
End Function

Private Function GetpathFromhandle(ByVal Handle As Long) As String

    Dim strbuffer As String
    strbuffer = Space$(32768)
    GetFinalPathnameByHandle Handle, strbuffer, Len(strbuffer), 0
    GetpathFromhandle = Replace$(Trim$(strbuffer), vbNullChar, "")


End Function
Private Function GetFinalPathnameByHandle(ByVal hFile As Long, ByRef lpszFilePath As String, ByVal cchFilePath As Long, ByVal dwFlags As FileNamefromHandleFlags) As Long
    GetFinalPathnameByHandle = GetFinalPathNameByHandleW(hFile, StrPtr(lpszFilePath), cchFilePath, dwFlags)
End Function
Public Function CloseHandle(ByVal hObject As Long) As Long
    'Debug.Print "Closing handle " & hObject
    CloseHandle = CloseHandleAPI(hObject)
End Function
'Private Declare Function FindFirstVolumeA Lib "kernel32.dll" (ByVal lpszVolumeName As String, ByVal cchBufferLength As Long) As Long
'Private Declare Function FindFirstVolumeMountPointA Lib "kernel32.dll" (ByVal lpszRootPathName As String, ByVal lpszVolumeMountPoint As String, ByVal cchBufferLength As Long) As Long
'Private Declare Function FindNextVolumeA Lib "kernel32.dll" (ByVal hFindVolume As Long, ByVal lpszVolumeName As String, ByVal cchBufferLength As Long) As Long
'Private Declare Function FindNextVolumeMountPointA Lib "kernel32.dll" (ByVal hFindVolumeMountPoint As Long, ByVal lpszVolumeMountPoint As String, ByVal cchBufferLength As Long) As Long
'Private Declare Function SetVolumeMountPointA Lib "kernel32.dll" (ByVal lpszVolumeMountPoint As String, ByVal lpszVolumeName As String) As Long
'Private Declare Function GetVolumeNameForVolumeMountPointA Lib "kernel32.dll" (ByVal lpszVolumeMountPoint As String, ByVal lpszVolumeName As String, ByVal cchBufferLength As Long) As Long
'Private Declare Function GetVolumePathNameA Lib "kernel32.dll" (ByVal lpszFileName As String, ByVal lpszVolumePathName As String, ByVal cchBufferLength As Long) As Long

Public Function FindFirstVolume(ByVal lpszVolumeName As String, ByVal cchBufferLength As Long) As Long

If MakeWideCalls Then
    FindFirstVolume = FindFirstVolumeW(StrPtr(lpszVolumeName), cchBufferLength)

Else
    FindFirstVolume = FindFirstVolumeA(lpszVolumeName, cchBufferLength)

End If


End Function

Public Function FindFirstVolumeMountPoint(ByVal lpszRootPathName As String, ByRef lpszVolumeMountPoint As String, ByVal cchBufferLength As Long) As Long
    '
    If MakeWideCalls Then
        FindFirstVolumeMountPoint = FindFirstVolumeMountPointW(StrPtr(lpszRootPathName), StrPtr(lpszVolumeMountPoint), cchBufferLength)
    Else
        FindFirstVolumeMountPoint = FindFirstVolumeMountPointA(lpszRootPathName, lpszVolumeMountPoint, cchBufferLength)
    End If
End Function
Public Function FindNextVolume(ByVal hFindVolume As Long, ByVal lpszVolumeName As String, ByVal cchBufferLength As Long) As Long
'
    If MakeWideCalls Then
        FindNextVolume = FindNextVolumeW(hFindVolume, StrPtr(lpszVolumeName), cchBufferLength)
    Else
        FindNextVolume = FindNextVolumeA(hFindVolume, lpszVolumeName, cchBufferLength)
    End If
    

End Function
Public Function FindNextVolumeMountPoint(ByVal hFindVolumeMountPoint As Long, ByRef lpszVolumeMountPoint As String, ByVal cchBufferLength As Long) As Long
'
End Function
Public Function SetVolumeMountPoint(ByVal lpszVolumeMountPoint As String, ByVal lpszVolumeName As String) As Long
'
If MakeWideCalls Then

    SetVolumeMountPoint = SetVolumeMountPointW(StrPtr(lpszVolumeMountPoint), StrPtr(lpszVolumeName))

Else
    SetVolumeMountPoint = SetVolumeMountPointA(lpszVolumeMountPoint, lpszVolumeName)


End If

End Function
Public Function GetVolumeNameForVolumeMountPoint(ByVal lpszVolumeMountPoint As String, ByRef lpszVolumeName As String, ByVal cchBufferLength As Long) As Long
'
If MakeWideCalls Then
    GetVolumeNameForVolumeMountPoint = GetVolumeNameForVolumeMountPointW(StrPtr(lpszVolumeMountPoint), StrPtr(lpszVolumeName), cchBufferLength)
Else
    GetVolumeNameForVolumeMountPoint = GetVolumeNameForVolumeMountPointA(lpszVolumeMountPoint, lpszVolumeName, cchBufferLength)
End If


End Function

Public Function SetCurrentDirectory(ByVal pszPath As String) As Long
    If MakeWideCalls Then

        SetCurrentDirectory = SetCurrentDirectoryW(StrPtr(pszPath))
    
    Else
        SetCurrentDirectory = SetCurrentDirectoryA(pszPath)
    
    End If
End Function
Public Function GetCurrentDirectory(ByRef strbuffer As String) As Long
    Dim retbuffer As String
    If MakeWideCalls Then
        retbuffer = Space$(32768)
        GetCurrentDirectory = GetCurrentDirectoryW(Len(retbuffer) - 1, StrPtr(retbuffer))
    Else
        retbuffer = Space$(MAX_PATH)
        GetCurrentDirectory = GetCurrentDirectoryA(Len(retbuffer) - 1, retbuffer)
    End If

    retbuffer = Trim$(Replace$(retbuffer, vbNullChar, ""))
    strbuffer = retbuffer
End Function
Public Function InternetOpen(ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszproxy As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long

    If MakeWideCalls Then
        InternetOpen = InternetOpenW(StrPtr(lpszAgent), dwAccessType, StrPtr(lpszproxy), StrPtr(lpszProxyBypass), dwFlags)
    Else
        InternetOpen = InternetOpenA(lpszAgent, dwAccessType, lpszproxy, lpszProxyBypass, dwFlags)
    
    End If
End Function
Public Function InternetConnect(ByRef hinternet As Long, ByVal lpszServerName As String, ByRef nServerPort As Long, ByVal lpszUserName As String, ByVal lpszPassword As String, ByVal dwService As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Long

    If MakeWideCalls Then
        InternetConnect = InternetConnectW(hinternet, StrPtr(lpszServerName), nServerPort, StrPtr(lpszUserName), StrPtr(lpszPassword), dwService, dwFlags, dwContext)
    Else
        InternetConnect = InternetConnectA(hinternet, lpszServerName, nServerPort, lpszUserName, lpszPassword, dwService, dwFlags, dwContext)
    End If
End Function



Public Function PathCompactPath(ByVal hDC As Long, ByVal pszPath As String, ByVal dx As Long) As Long
'pathCompactPathA
If MakeWideCalls() Then
    PathCompactPath = PathCompactPathW(hDC, StrPtr(pszPath), dx)



Else
    PathCompactPath = PathCompactPathA(hDC, pszPath, dx)


End If


pszPath = Mid$(pszPath, 1, InStr(1, pszPath, vbNullChar, vbBinaryCompare))

End Function
Public Function PathCompactPathEx(ByRef pszOut As String, ByVal pszSrc As String, ByVal cchmax As Long, ByVal dwFlags As Long) As Long
    If MakeWideCalls Then
        'Unicode.
        PathCompactPathEx = PathCompactPathExW(StrPtr(pszOut), StrPtr(pszSrc), cchmax, dwFlags)
    
    Else
        'ANSI
        PathCompactPathEx = PathCompactPathExA(pszOut, pszSrc, cchmax, dwFlags)
    
    End If

End Function
Public Function PathIsSameRoot(pszPath1 As String, pszPath2 As String) As Boolean

    If MakeWideCalls() Then
        PathIsSameRoot = PathIsSameRootW(StrPtr(pszPath1), StrPtr(pszPath2))
    Else
        PathIsSameRoot = PathIsSameRootA(pszPath1, pszPath2)
    End If
End Function
Public Function TranslateAPIErrorCode(ByVal ApiError As Long) As Long
    'Translates a Windows API Error Code to a Visual Basic Error.
    
    Static mLookup As Scripting.Dictionary
    If mLookup Is Nothing Then
        Set mLookup = New Dictionary
        'add values.
        
        
        
    End If
    
    
    



End Function

Public Function GetFileAttributes(ByVal StrFilename As String) As Long
    Dim w32 As WIN32_FIND_DATA
    Dim hfind As Long
    hfind = FindFirstFile(StrFilename, w32)
    If Replace$(Trim$(w32.cFileName), vbNullChar, "") <> "" Then
        GetFileAttributes = w32.dwFileAttributes
    Else
    If MakeWideCalls Then
        If Not IsUNCPath(StrFilename) Then
            StrFilename = "//?/" & StrFilename
        End If
        GetFileAttributes = GetFileAttributesW(StrPtr(StrFilename))
    Else
        GetFileAttributes = GetFileAttributesA(StrFilename)

    End If
       
    End If


End Function
'Public Function GetFileAttributes(ByVal StrFilename As String) As Long
'    Dim rt As Long
'    If MakeWideCalls Then
'        If Not IsUNCPath(StrFilename) Then
'            StrFilename = "//?/" & StrFilename
'        End If
'        ret = GetFileAttributesW(StrPtr(StrFilename))
'    Else
'        ret = GetFileAttributesA(StrFilename)
'
'    End If
'
'    If ret = -1 Then
'        'double-check...
'        Dim w32 As WIN32_FIND_DATA
'        w32.dwFileAttributes
'
'    End If
'
'
'    GetFileAttributes = ret
'
'End Function
''VOID CALLBACK FileIOCompletionRoutine(
'  __in  DWORD dwErrorCode,
'  __in  DWORD dwNumberOfBytesTransfered,
'  __in  LPOVERLAPPED lpOverlapped
');

'routines used to add/remove streams that are waiting for a FileIOCompletionRoutine.
Public Function AddAsyncStream(ByVal obj As Object) As Long
    'adds to the collection. returns index (current count)
    
    If mWaitingAsync Is Nothing Then
        Set mWaitingAsync = New Collection
    End If
    mWaitingAsync.Add obj
    AddAsyncStream = mWaitingAsync.count + 1
    
    
End Function
Public Function removeAsyncStream(obj As Long) As Long
    mWaitingAsync.Remove obj



End Function
Public Function GetAsyncStream(ObjIndex As Long) As Object
    Set GetAsyncStream = mWaitingAsync.Item(ObjIndex)
End Function
Public Sub FileIOCompletionRoutine(ByVal dwErrorCode As Long, ByVal dwBytesTransferred As Long, OverlappedVar As OVERLAPPED)
'since we will only be using this in the context of WriteFileEx and ReadFileEx, we can use the "hevent" member to store the index into our array of pending file operations.
    Dim getobj As Object
    Dim casted As IAsyncProcess
    Debug.Print "FileIOCompletionRoutine"
    Set getobj = GetAsyncStream(OverlappedVar.hEvent)
    Set casted = getobj
    'casted.ExecAsync
    






End Sub



Public Function GetErrorMode() As SetErrorModeConstants
    Dim tmp As Long
    tmp = SetErrorMode(0)
    GetErrorMode = tmp
    SetErrorMode tmp
    
End Function
Public Function Random() As String

Const contextname As String = "Microsoft Enhanced Cryptographic Provider v1.0"



End Function



'Public Function GetFriendlyEXEName(ByVal StrEXE As String) As String
'Dim lpData As Long
'Dim lpSize As Long, stralloc As String
'Dim lpreturnstr As String, lpretlen As Long
'
'lpSize = GetFileVersionInfoSize(StrEXE, lpData)
'Dim ret As Long
'stralloc = Space$(lpSize)
'lpData = VarPtr(lpSize)
'ret = GetFileVersionInfo(StrEXE, 0, lpSize, ByVal lpData)
'lpreturnstr = Space$(255)
'VerQueryValue ByVal lpData, "FileDescription" & vbNullChar, lpreturnstr, lpretlen
'GetFriendlyEXEName = lpreturnstr
'
'End Function



'API WRAPPERS:
Public Sub SHEmptyRecycleBin(ByVal hwnd As Long, ByVal pszRootPath As Long, ByVal dwFlags As Long)
    If MakeWideCalls Then
        SHEmptyRecycleBinW hwnd, StrPtr(pszRootPath), dwFlags
    Else
        SHEmptyRecycleBinA hwnd, pszRootPath, dwFlags
    End If



End Sub


Public Function GetCompressedFileSize(ByVal lpFileName As String, ByRef lpFileSizeHigh As Long) As Long
    
    If MakeWideCalls Then
        GetCompressedFileSize = GetCompressedFileSizeW(StrPtr(lpFileName), lpFileSizeHigh)
    Else
        GetCompressedFileSize = GetCompressedFileSizeA(lpFileName, lpFileSizeHigh)
    End If
    
End Function
Public Function SetFileAttributes(ByVal strFile As String, ByVal Attributes As FileAttributeConstants)

If MakeWideCalls Then
     If Not IsUNCPath(strFile) Then
            strFile = "//?/" & strFile
        End If
    SetFileAttributesW StrPtr(strFile), Attributes
Else
    SetFileAttributesA strFile, Attributes
End If

End Function
Public Sub SetFileTimes(ByVal mvarFileName As String, ByVal DateCreated As Date, ByVal lastAccess As Date, ByVal LastWrite As Date)
    Dim hFile As Long
    
    Dim ftdatecreated As FILETIME, ftlastaccess As FILETIME, ftlastwrite As FILETIME
    
    ftdatecreated = Date2FILETIME(DateCreated)
    ftlastaccess = Date2FILETIME(lastAccess)
    ftlastwrite = Date2FILETIME(LastWrite)
    'step one: open the file.
    hFile = CreateFile(mvarFileName, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, 0, 0)
    
    If hFile > 0 Then
        'Success!
        
        
        SetFileTime hFile, ftdatecreated, ftlastaccess, ftlastwrite
        
        
        CloseHandle hFile
    Else
    
    
        RaiseAPIError Err.LastDllError, "MdlFileSystem::SetFileTimes()"
    
    End If



End Sub

'Private Const MAX_PATH = 260
'Private Declare Function CreateDirectoryA Lib "kernel32.dll" (ByVal lpPathName As String, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
'Private Declare Function RemoveDirectoryA Lib "kernel32.dll" (ByVal lpPathName As String) As Long
Public Function CreateDirectoryEx(ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, ByVal lpSecurityAttributes As Long) As Long
If MakeWideCalls Then
    CreateDirectoryEx = CreateDirectoryExW(StrPtr(lpTemplateDirectory), StrPtr(lpNewDirectory), lpSecurityAttributes)
Else
    CreateDirectoryEx = CreateDirectoryExA(lpTemplateDirectory, lpNewDirectory, lpSecurityAttributes)
End If


End Function

Public Function CreateDirectory(ByVal lpPathName As String, ByVal lpSecurityAttributes As Long) As Long

If MakeWideCalls Then
    CreateDirectory = CreateDirectoryW(StrPtr(lpPathName), lpSecurityAttributes)
Else
    CreateDirectory = CreateDirectoryA(lpPathName, lpSecurityAttributes)
End If

End Function
Public Function RemoveDirectory(ByVal lpPathName As String) As Long

If MakeWideCalls Then
    RemoveDirectory = RemoveDirectoryW(StrPtr(lpPathName))
Else
    RemoveDirectory = RemoveDirectoryA(lpPathName)
End If

End Function

Public Function SHGetSpecialFolderPath(ByVal hwnd As Long, ByRef folderlocation As String, FolderConst As SpecialFolderConstants, ByVal fCreate As Long) As Long
    
'    If MakeWideCalls Then
'        FolderLocation = Space$(MAX_PATH * 2)
'        SHGetSpecialFolderPath = SHGetSpecialFolderPathW(hwnd, StrPtr(FolderLocation), FolderConst, fCreate)
'        'folderlocation
'    Else
        folderlocation = Space$(MAX_PATH)
        SHGetSpecialFolderPath = SHGetSpecialFolderPathA(hwnd, folderlocation, FolderConst, fCreate)
'    End If




End Function

Public Function ListStreams_Vista(ByVal Spath As String) As CAlternateStreams
    Dim newstreams As CAlternateStreams
    Dim createalt As CAlternateStream
    Dim useBuff As WIN32_FIND_STREAM_DATA
    Dim FHandle As Long, usename As String
    Set newstreams = New CAlternateStreams
    FHandle = FindFirstStreamW(StrPtr(Spath), 0, useBuff, 0)
    Do
        Set createalt = New CAlternateStream
        usename = Replace$(Mid$(Trim$(StrConv(useBuff.cStreamName, vbFromUnicode)), 2), vbNullChar, "")
        'if the stream name ends with ":$DATA" then strip that off... but only if the name is longer  then 7 characters...
        
        If Len(usename) > 7 Then
            If Right$(usename, 6) = ":$DATA" Then
            usename = Mid$(usename, 1, Len(usename) - 6)
            
            End If
        
        End If
        
        If usename <> "" Then
            With createalt
                .Init Spath, usename, useBuff.StreamSize.lowpart, useBuff.StreamSize.highpart, 0
            End With
            newstreams.Add createalt
        End If
    
        ZeroMemory useBuff, Len(useBuff)
        If FindNextStreamW(FHandle, useBuff) = 0 Then
            Exit Do
        End If
        
    
    Loop
    FindClose FHandle

    Set ListStreams_Vista = newstreams
End Function
Public Function HardLinks_Vista(ByVal Spath As String, Optional ByRef count As Long) As String()
'returns a list of files with hardlinks to this one.





End Function
Public Function HardLinks_PreVista(ByVal Spath As String, Optional ByRef count As Long) As String()
    'Awful... just... awful...
    
    
    'anyway.... get the file index, and then literally search on all files on that drive for the links...
    'returns all hardlinks OTHER then this file.
    
    Dim FolderStack() As Directory, StackTop As Long
    Dim gotfile As CFile
    Dim findex As Double
    Dim ret() As String
    Dim linkcount As Long
    Dim startfolder As Directory
    Dim currfile As CFile
    Dim currdir As Directory
    Set gotfile = FileSystem.GetFile(Spath)
    'retrieve the file index...
    findex = gotfile.FileIndex
    If gotfile.HardLinkCount = 1 Then
        'only one link... this one.
        count = 0
        
    Else
        Set startfolder = gotfile.Directory.Volume.RootFolder
        'find all files, then all folders.
        StackTop = 1
        ReDim FolderStack(1 To 1)
        Set FolderStack(1) = startfolder
        Do Until StackTop = 0
            'Grab the topmost item....
            Dim topItem As Directory
            Set topItem = FolderStack(StackTop)
            StackTop = StackTop - 1
            If StackTop > 0 Then
            ReDim Preserve FolderStack(1 To StackTop)
            End If
            '"remove" this item...
            'If StrComp(Left$(topItem.Path, 9), "D:\VBPROJ", vbTextCompare) = 0 Then Stop
            Dim LoopFile As CFile, LoopDir As Directory
            'now, loop through all files...
            With topItem.Files.GetWalker
            Do Until .GetNext(LoopFile) Is Nothing
            If StrComp(LoopFile.Fullpath, "D:\vbproj\vb\testhl.txt", vbTextCompare) = 0 Then Stop
                If LoopFile.FileIndex = findex Then
                    'add to our return array...
                    If LoopFile.Fullpath <> gotfile.Fullpath Then
                        linkcount = linkcount + 1
                        ReDim Preserve ret(1 To linkcount)
                        ret(linkcount) = LoopFile.Fullpath
                        If linkcount = (gotfile.HardLinkCount - 1) Then
                            'all links found...
                            'break out...
                            HardLinks_PreVista = ret
                            count = linkcount
                            Exit Function
                        End If
                    End If
                   ' FindFirstStreamW
                End If
            Loop
            End With
            'OK, now loop through the directories...
            With topItem.Directories.GetWalker
                Do Until .GetNext(LoopDir) Is Nothing
                    'push it into the stack...
                    StackTop = StackTop + 1
                    ReDim Preserve FolderStack(1 To StackTop)
                    Set FolderStack(StackTop) = LoopDir
                    'If InStr(1, LoopDir.Path, "vbproj", vbTextCompare) > 0 Then Stop
                    Debug.Print "stacktop=" & StackTop & " Folder " & LoopDir.Path
                Loop
            
            End With
            
        
        Loop
    
    End If


End Function
Public Function ListStreams(Spath As String) As CAlternateStreams
    'Purpose: call ListStreams_Vista on Vista- Call ListStreams_NT for other OS's.
    
    'Non-NT platforms don't have Streams...
    If Not IsWinNT Then
        Set ListStreams = Nothing
        Exit Function
    End If
    
    
    If IsVistaOrLater Then
        Set ListStreams = ListStreams_Vista(Spath)
    Else
        Set ListStreams = ListStreams_NT(Spath)
    End If



End Function
Public Function ListStreams_NT(Spath As String) As CAlternateStreams
    Dim IOS As IO_STATUS_BLOCK
    Dim BBuf() As Byte
    Dim FSInfo As FILE_STREAM_INFORMATION
    Dim LBuf As Long, LInfo As Long, lRet As Long
    Dim lErr As Long
    Dim sName As String, SNames As String
    
    Dim newstreams As CAlternateStreams
    Set newstreams = New CAlternateStreams
    
    newstreams.Owner = Spath
    On Error Resume Next
    'ListStreams = ""
    lRet = CreateFile(Spath, DesiredAccessFlags.STANDARD_RIGHTS_READ, FILE_SHARE_READ, 0&, OPEN_EXISTING, _
    FILE_FLAG_BACKUP_SEMANTICS, 0&)
    If (lRet = -1) Then Exit Function
    
    LBuf = 4096
    lErr = 234
    ReDim BBuf(1 To LBuf)
    
    Do While lErr = 234
    
        lErr = NtQueryInformationFile(lRet, IOS, ByVal VarPtr(BBuf(1)), LBuf, _
        ByVal FileStreamInformation)
        If (lErr = 234) Then
        LBuf = LBuf + 4096
        ReDim BBuf(1 To LBuf)
        End If
        
    Loop
    
    LInfo = VarPtr(BBuf(1))
    Dim newstream As CAlternateStream
    Do
    
        CopyMemory ByVal VarPtr(FSInfo.NextEntryOffset), ByVal LInfo, Len(FSInfo)
        'CopyMemory ByVal VarPtr(FSInfo.StreamName(0)), ByVal LInfo + 24, _
'
        'FSInfo.StreamNameLength
        sName = Left$(FSInfo.StreamName, FSInfo.StreamNameLength / 2)
        
        If (InStr(1, sName, DATA_1, 1) = 0) And (InStr(1, sName, DATA_2, 1) = 0) _
        And (sName <> "") Then
            'SNames = SNames & Mid$(sName, 2, Len(sName) - 7) & " * " & _
            CStr(FSInfo.StreamSize) & "|"
          
            Set newstream = New CAlternateStream
            newstream.Init Spath, Mid$(sName, 2, Len(sName) - 7), _
            FSInfo.StreamSize, FSInfo.StreamSizeHi
            newstreams.Add newstream
            
        End If
        
        If FSInfo.NextEntryOffset Then
            LInfo = LInfo + FSInfo.NextEntryOffset
        Else
            Exit Do
        End If
        
    Loop
    CloseHandle lRet
    If (Len(SNames) > 0) Then SNames = Left$(SNames, (Len(SNames) - 1))
    Set ListStreams_NT = newstreams
    'Stop
End Function






Public Function GetDriveForNtDeviceName(ByVal sDeviceName As String) As String
Dim sFoundDrive As String
Dim strdrives As String
Dim DriveStr() As String
Dim vDrive As String, I As Long, ret As Long
strdrives = Space$(256)
ret = GetLogicalDriveStrings(255, strdrives)
strdrives = Trim$(Replace$(strdrives, vbNullChar, " "))
DriveStr = Split(strdrives, " ")
   'For Each vDrive In GetDrives()
   For I = 0 To UBound(DriveStr)
    vDrive = DriveStr(I)
      If StrComp(GetNtDeviceNameForDrive(vDrive), sDeviceName, vbTextCompare) = 0 Then
         sFoundDrive = vDrive
         Exit For
      End If
   Next I
   
   GetDriveForNtDeviceName = sFoundDrive
   
End Function

Public Function GetNtDeviceNameForDrive( _
   ByVal sDrive As String) As String
Dim bDrive() As Byte
Dim bresult() As Byte
Dim lR As Long
Dim sDeviceName As String

   If Right(sDrive, 1) = "\" Then
      If Len(sDrive) > 1 Then
         sDrive = Left(sDrive, Len(sDrive) - 1)
      End If
   End If
   bDrive = sDrive
   ReDim Preserve bDrive(0 To UBound(bDrive) + 2) As Byte
   ReDim bresult(0 To MAX_PATH * 2 + 1) As Byte
   lR = QueryDosDeviceW(VarPtr(bDrive(0)), VarPtr(bresult(0)), MAX_PATH)
   If (lR > 2) Then
      sDeviceName = bresult
      sDeviceName = Left(sDeviceName, lR - 2)
      GetNtDeviceNameForDrive = sDeviceName
   End If
   
End Function



Public Function GetSpecialFolder(hwnd As Long, FolderConst As SpecialFolderConstants) As String

Dim folderlocation As String, ret As Long
folderlocation = Space$(2048)
ret = SHGetSpecialFolderPath(hwnd, folderlocation, FolderConst, 0)

GetSpecialFolder = Trim$(Replace$(folderlocation, vbNullChar, ""))



End Function
Public Function GetSpecialFolderPidl(ByVal hwnd As Long, ByVal folder As SpecialFolderConstants) As Long


Dim ret As Long

SHGetSpecialFolderLocation hwnd, folder, ret
GetSpecialFolderPidl = ret


End Function
Public Function MakeWideCalls() As Boolean

  'MakeWideCalls = (m_IsWinNt And (m_WideCallSupport <> AnsiVersion))
  MakeWideCalls = IsWinNT And Not ForceANSI
  
End Function

Public Function SizeOfString() As Long
  If MakeWideCalls Then
    SizeOfString = 2
  Else
    SizeOfString = 1
  End If
End Function

'ANSI/WIDE WRAPPERS


Public Function CreateFileMapping(ByVal hFile As Long, ByVal lpFileMappingAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As Variant) As Long
    Dim lpstrname As String
    If Not IsNull(lpName) Then
    lpstrname = str(lpName)
    End If
    If MakeWideCalls Then
        If IsNull(lpName) Or IsEmpty(lpName) Or lpName = vbNullString Then
        CreateFileMapping = CreateFileMappingW(hFile, lpFileMappingAttributes, flProtect, dwMaximumSizeHigh, dwMaximumSizeLow, ByVal 0&)
        Else
            CreateFileMapping = CreateFileMappingW(hFile, lpFileMappingAttributes, flProtect, dwMaximumSizeHigh, dwMaximumSizeLow, StrPtr(lpstrname))
        End If
    Else
        If IsNull(lpName) Or IsEmpty(lpName) Or lpName = vbNullString Then
            CreateFileMapping = CreateFileMappingALong(hFile, lpFileMappingAttributes, flProtect, dwMaximumSizeHigh, dwMaximumSizeLow, ByVal 0&)
        Else
            CreateFileMapping = CreateFileMappingA(hFile, lpFileMappingAttributes, flProtect, dwMaximumSizeHigh, dwMaximumSizeLow, lpstrname)
        End If
    End If

End Function
Private Function GetDefaultSecAttr() As SECURITY_ATTRIBUTES
    Dim ret As SECURITY_ATTRIBUTES
    ret.bInheritHandle = True
    ret.lpSecurityDescriptor = 0
    GetDefaultSecAttr = ret

End Function
Public Function CreateFile(ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Dim lpfileret As Long

Dim SecAttrUse As SECURITY_ATTRIBUTES

SecAttrUse = GetDefaultSecAttr
If Not IsUNCPath(lpFileName) Then
If Left$(lpFileName, 2) <> "\\" Then
lpFileName = "\\?\" & lpFileName
End If
End If

    If MakeWideCalls() Then
        lpfileret = CreateFileW(StrPtr(lpFileName), dwDesiredAccess, dwShareMode, SecAttrUse, dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile)
    
        
    Else
        lpfileret = CreateFileA(lpFileName, dwDesiredAccess, dwShareMode, SecAttrUse, dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile)
    
    End If
    'Debug.Print "CreateFile returning handle " & lpfileret & " for file  " & lpFileName
    CreateFile = lpfileret
    If lpfileret = -1 Then
        Debug.Print "Error accessing " & lpFileName & ":" & GetAPIErrStr(Err.LastDllError)
    
    End If


End Function











Sub Main()
    Set Winmetrics = New SystemMetrics
    Set LargeIcons = New cVBALImageList
    LargeIcons.ColourDepth = ILC_COLOR32
    SmallIcons.ColourDepth = ILC_COLOR32
    LargeIcons.IconSizeY = Winmetrics.LargeIconSize
    LargeIcons.IconSizeX = Winmetrics.LargeIconSize
    LargeIcons.Create
    Set SmallIcons = New cVBALImageList
    SmallIcons.IconSizeX = Winmetrics.SmallIconSize
    SmallIcons.IconSizeX = Winmetrics.SmallIconSize
    SmallIcons.Create
    Set ShellIcons = New cVBALImageList
    
    Dim quickdllstream As FileStream
    Dim resbytes() As Byte, quickdllpath
    'expand DLLs needed.
    quickdllpath = FileSystem.GetSpecialFolder(CSIDL_SYSTEMX86).Path & "quick32.dll"
    
    If Not FileSystem.Exists(quickdllpath) Then
    resbytes = LoadResData("QUICK", "DLL")
    Set quickdllstream = FileSystem.CreateStream(quickdllpath)
    quickdllstream.WriteBytes resbytes
    quickdllstream.CloseStream
    End If
'    Dim TESTIT As String
'    Dim pidlRel As Long, pidlfile As Long
'    Dim relfolder As olelib.IShellFolder
    InitKnownFolders
End Sub
Public Sub DBL2LI(ByVal Dbl As Double, ByRef lopart As Long, ByRef hipart As Long)
    Dim mungec As MungeCurr
    Dim mungeli As MungeLong
    mungec.CurrA = (Dbl / 10000#)
    LSet mungeli = mungec
    lopart = mungeli.LongA
    hipart = mungeli.LongB
End Sub

Public Function LI2DBL(lopart As Long, hipart As Long) As Double
    Dim mungel As MungeLong
    Dim mungec As MungeCurr
    mungel.LongA = lopart
    mungel.LongB = hipart
    LSet mungec = mungel
    LI2DBL = mungec.CurrA * 10000#




End Function
Public Function FileTime2Date(ftime As FILETIME) As Date
    Dim sTime As SYSTEMTIME
    Dim createdate As Date
    FileTimeToSystemTime ftime, sTime
    With sTime
    createdate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    
    End With
    
    
    FileTime2Date = createdate
    
    
End Function
Public Function Date2FILETIME(Dateuse As Date) As FILETIME
    Dim retStruct As FILETIME
    Dim FromSTime As SYSTEMTIME
    With FromSTime
    .wSecond = Second(Dateuse)
    .wMinute = Minute(Dateuse)
    .wHour = Hour(Dateuse)
    .wDay = Day(Dateuse)
    
    .wMonth = Month(Dateuse)
    .wYear = Year(Dateuse)
    .wDayOfWeek = Weekday(Dateuse)
    
    End With
    SystemTimeToFileTime FromSTime, retStruct
    Date2FILETIME = retStruct
        
    
    
End Function
Public Sub AppendSlash(ByRef Path As String)
    If (Right$(Path, 1) <> "\" And Right$(Path, 1) <> "/") Then Path = Path & "\"

End Sub
Public Function FixPath(ByVal Path As String) As String
    'reformats a path to use understandable slashes, and other things.
    Dim Startreplace As Long
    'proper UNC form is
    '//SERVER/SHARE
    If IsUNCPath(Path) Then
        'Startreplace = InStr(3, Path, "/") + 1
        FixPath = Replace$(Path, "\", "/")
    Else
        FixPath = Replace$(Path, "/", "\")
        'Startreplace = 1
        
    
    End If
    


End Function
Public Sub RaiseAPIError(ByVal ErrCode As Long, ByVal ErrSource As String)

Dim MessageStr As String
MessageStr = GetAPIError(ErrCode)
If MessageStr = "" Then MessageStr = "Unexpected Error in " & ErrSource

Err.Raise BCFileErrorBase + ErrCode, ErrSource, MessageStr



End Sub
Public Function GetAPIError(ByVal ErrCode As Long) As String
    'raises a Windows API error.
 '   FormatMessage(
 ' FORMAT_MESSAGE_ALLOCATE_BUFFER |
 ' FORMAT_MESSAGE_FROM_SYSTEM,
 ' NULL,
 ' Err.LastDLLError(),
 ' MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), //The user default language
 ' (LPTSTR) &lpMessageBuffer,
 ' 0,
 ' NULL );
Dim lpBuffer As String
lpBuffer = Space$(128)
FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, ErrCode, 0, lpBuffer, Len(lpBuffer), ByVal 0&
GetAPIError = Replace$(Trim$(lpBuffer), vbNullChar, "")


End Function
Public Function ParseVolume(ByVal StrFrom As String) As String

Dim ret As String
ParsePathParts StrFrom, ret, , , , , Parse_Volume
ParseVolume = ret


End Function
Public Function ParseStreamName(ByVal StrFrom As String) As String
    Dim retstr As String
    ParsePathParts StrFrom, , , , , retstr, Parse_Stream
    
    ParseStreamName = retstr



End Function
Public Function ParseExtension(ByVal StrFrom As String) As String
    Dim ext As String
    ParsePathParts StrFrom, , , , ext
    ParseExtension = ext
End Function
Public Function ParseFilename(ByVal StrFrom As String, Optional ByVal WithExtension As Boolean = True) As String


    Dim fname As String, Fextension As String
    ParsePathParts StrFrom, , , fname, Fextension
    If WithExtension Then
        ParseFilename = fname & "." & Fextension
    Else
        ParseFilename = fname
    End If



End Function
Public Function ParsePath(ByVal StrFrom As String) As String
    'parse the path from a specification
    Dim retval As String
    ParsePathParts StrFrom, , retval
    
    ParsePath = retval


End Function

'Public Function TrimNull$(ByVal szString As String)
'    TrimNull = Trim$(Replace$(szString, vbNullChar, ""))
'End Function

'SH stuff.
Public Function DisplayName(ByVal lpszPath As String, Optional ByVal AssumeExist As Boolean) As String
    Dim finfo As SHFILEINFO
    Dim ValFlag As ShellFileInfoConstants
    ValFlag = SHGFI_DISPLAYNAME
    FixAssume ValFlag, AssumeExist
    SHGetFileInfo lpszPath, 0, finfo, Len(finfo), SHGFI_DISPLAYNAME
    DisplayName = TrimNull$(finfo.szDisplayName)
    


End Function
Public Function FileTypeName(ByVal strFile As String, Optional ByVal AssumeExist As Boolean) As String
    Dim finfo As SHFILEINFO
    Dim ValFlag As ShellFileInfoConstants
    ValFlag = SHGFI_TYPENAME
    FixAssume ValFlag, AssumeExist
    SHGetFileInfo strFile, 0, finfo, Len(finfo), SHGFI_TYPENAME
    FileTypeName = TrimNull$(finfo.szTypeName)
End Function








Private Sub FixAssume(ByRef ValFix As ShellFileInfoConstants, ByVal Assume As Boolean)
    If Assume Then
        ValFix = ValFix Or SHGFI_USEFILEATTRIBUTES
    End If
End Sub
Function PidlFromPath(Spath As String) As Long
    Dim pidl As Long, F As Long
    F = SHGetPathFromIDList(pidl, Spath)
    If F Then PidlFromPath = pidl
End Function
Public Function ShowExplorerMenu(ByVal hWndOwner As Long, ByVal pszPath As String, Optional X As Long = -1, Optional Y As Long = -1, Optional menucallback As IContextCallback = Nothing, Optional CMFFlags As QueryContextMenuFlags = CMF_EXPLORE) As Long
    'displays the Explorer menu.
    
    'MFoldTool.ContextPopMenu hwndOwner, pszPath, x, y
    
    
    Dim pidlRel As Long, pidlpath As Long, pidlfile As Long
    Dim deskfolder As olelib.IShellFolder
    Dim parentfolder As olelib.IShellFolder
    Dim Pointuse As POINTAPI
    'Always relative to desktop.
    SHGetDesktopFolder deskfolder

    'PidlPath = SHSimpleIDListFromPath(pszPath)
    'PidlPath = PidlFromPath(pszPath)
     If X = -1 And Y = -1 Then
            GetCursorPos Pointuse
        Else
            Pointuse.X = X
            Pointuse.Y = Y
        End If
    
    If Len(pszPath) <= 3 Then
        
        'why, it's a drive spec.
        deskfolder.ParseDisplayName hWndOwner, 0, StrPtr(pszPath), 0, pidlfile, 0
        Set parentfolder = deskfolder
    
    Else
    
        Set parentfolder = FolderFromItem(hWndOwner, pszPath, pidlpath)
       
        'Call ShowShellContextMenu(hwndOwner, DeskFolder, 1, 0, Pointuse)
        'Call ShowShellContextMenu(hwndOwner, DeskFolder, 1, PidlPath, Pointuse)
        
        
        'Current Work: this line errs when a Drive name is specified.
        Dim fnamestart As Long
        fnamestart = InStrRev(pszPath, "\") + 1
        If fnamestart = 0 Then
            fnamestart = InStrRev(pszPath, "/") + 1
        End If
        
        
        parentfolder.ParseDisplayName hWndOwner, 0, StrPtr(Mid$(pszPath, fnamestart)), 0, pidlfile, 0
        
    End If
    
    'To allow for multiple files:
    
    'they would all need to be in the same dir, BTW-
    
    'retrieve common parentfolder
    'acquire pidl for each Item for which a context menu is to be shown-
    '(ParseDisplayName method)
    
    'with the new array of Pidls, call the ShowShellContextMenu function, with an appropriate count and the first item in the array.
    ShowShellContextMenu hWndOwner, parentfolder, 1, pidlfile, Pointuse, menucallback, , CMFFlags
    'Call ShowShellContextMenu(hwndOwner, PidlPath, 1, 0, Pointuse)





End Function

Public Function ShowExplorerMenuMulti(ByVal hWndOwner As Long, ByVal pszPath As String, StrFiles() As String, Optional X As Long = -1, Optional Y As Long = -1, Optional CallbackObject As IContextCallback = Nothing) As Long
    'displays the Explorer menu.
    
    'MFoldTool.ContextPopMenu hwndOwner, pszPath, x, y
    
    
    Dim pidlRel As Long, pidlpath As Long, pidlfile() As Long
    Dim deskfolder As olelib.IShellFolder
    Dim parentfolder As olelib.IShellFolder
    Dim Pointuse As POINTAPI, I As Long
    'Always relative to desktop.
    SHGetDesktopFolder deskfolder

    'PidlPath = SHSimpleIDListFromPath(pszPath)
    'PidlPath = PidlFromPath(pszPath)
     If X = -1 And Y = -1 Then
            GetCursorPos Pointuse
        Else
            Pointuse.X = X
            Pointuse.Y = Y
        End If
    
    'If Len(pszPath) <= 3 Then
        
        'why, it's a drive spec.
        'DeskFolder.ParseDisplayName hwndOwner, 0, StrPtr(pszPath), 0, pidlfile, 0
   '     Set ParentFolder = DeskFolder
    'Else
    'Else
    
        Set parentfolder = FolderFromItem(hWndOwner, pszPath, pidlpath)
    'End If
        'Call ShowShellContextMenu(hwndOwner, DeskFolder, 1, 0, Pointuse)
        'Call ShowShellContextMenu(hwndOwner, DeskFolder, 1, PidlPath, Pointuse)
        
        
        'Current Work: this line errs when a Drive name is specified.
          'retrieve common parentfolder
        'acquire pidl for each Item for which a context menu is to be shown-
        '(ParseDisplayName method)
        'with the new array of Pidls, call the ShowShellContextMenu function, with an appropriate count and the first item in the array.
        ReDim pidlfile(UBound(StrFiles))
        For I = 0 To UBound(StrFiles)
            parentfolder.ParseDisplayName hWndOwner, 0, StrPtr(pszPath & StrFiles(I)), 0, pidlfile(I), 0
        Next
        'ParentFolder.ParseDisplayName hwndOwner, 0, StrPtr(Mid$(pszPath, InStrRev(pszPath, "\") + 1)), 0, pidlfile, 0
        
   ' End If
    
    'To allow for multiple files:
    
    'they would all need to be in the same dir, BTW-
    
  
    
    
    ShowShellContextMenu hWndOwner, parentfolder, UBound(pidlfile) + 1, pidlfile(0), Pointuse, CallbackObject
    'Call ShowShellContextMenu(hwndOwner, PidlPath, 1, 0, Pointuse)





End Function
Public Function IsFileName(ByVal Spec As String) As Boolean
    'returns wether Spec specifies a Filename.
    'for example:
    
    'C:\ would return false.
    'C:\a would return true, unless a folder currently exists in C called "a"
    
    Dim Attribs As FileAttributeConstants
    Attribs = GetFileAttributes(Spec)
    
    If Attribs > 0 Then
        If (Attribs And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Or _
        (Attribs And FILE_ATTRIBUTE_DEVICE) = FILE_ATTRIBUTE_DEVICE Then
            IsFileName = False
        Else
        
            IsFileName = True
        End If
    
    Else
        IsFileName = False 'not found...
    End If



End Function
Public Function isDirectory(ByVal Spec As String) As Boolean
    Dim Attribs As FileAttributeConstants
    Attribs = GetFileAttributes(Spec)
    isDirectory = ((Attribs And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
End Function
Public Function GetName(ByVal Path As String) As String
    'returns the Name portion of the path.
    'for example:
    
    '"TEST.DLL"
    'would return TEST
    'and C:\Windows\System32.old\
    'should return System32
    
    
End Function
Public Function GetStream(ByVal OfPath As String) As String
    'returns the Stream name of the file.
    'look backwards through the path, if a ":" is closer then a \, then grab the value between the colon and the end of the string.
    Dim LastColon As Long, lastSlash As Long
    OfPath = Replace$(OfPath, "/", "\")
    LastColon = InStrRev(OfPath, ":")
    lastSlash = InStrRev(OfPath, "\")
    If LastColon > lastSlash Then
        'stream name present....
        GetStream = Mid$(OfPath, LastColon)
    Else
        'no stream name.
        GetStream = ""
        End If
    
    
End Function
Public Function GetDrive(ByVal Path As String) As CVolume
    Dim retme As CVolume
    Set retme = New CVolume
    retme.Init Path
    Set GetDrive = retme
End Function
'Code to extract portions from a file path.
Public Function GetVolume(ByVal Path As String) As String
'Precondition: Path is an absolute path.

'supported "Volume" syntaxes:

'<Driveletter>:

'//VolumeName/
Path = Replace$(Path, "/", "\")
If Left$(Path, 2) = "\\" Then

'UNC
    GetVolume = Replace$(Mid$(Path, 1, InStr(3, Path, "\")), "\", "/")
ElseIf Mid$(Path, 2, 1) = ":" Then
    'Drive spec.
    GetVolume = Mid$(Path, 1, 3)
End If


End Function


Public Sub CopyFile(ByVal ExistingFilename As String, ByVal newFilename As String, Progressobj As IProgressCallback, Optional ByVal Flags As CopyFileExFlags)
    Dim retval As Long, brcancel As Long, copyflags As Long
    Dim Cookie As Long
    Cookie = CLng(AddCPCallback(Progressobj))
    retval = CopyFileEx(StrPtr(ExistingFilename), StrPtr(newFilename), _
    AddressOf CopyProgressRoutine, ByVal Cookie, 0, 2)
    RemoveCPCallback Cookie

End Sub
Public Sub MoveFile(ByVal ExistingFilename As String, ByVal Destination As String, Progressobj As IProgressCallback)

End Sub




' Display a context menu from a folder
' Based on C code by Jeff Procise in PC Magazine
' Destroys any pidl passed to it, so pass duplicate if necessary
'Function ContextPopMenu(ByVal hwnd As Long, vItem As Variant, _
'                        ByVal x As Long, ByVal y As Long) As Boolean
'    InitIf  ' Initialize if in standard modue
'
'    Dim folder As IShellFolder, pidlMenu As Long
'    Dim menu As IContextMenu3, ici As CMINVOKECOMMANDINFO
'    Dim iCmd As Long, f As Boolean, hMenu As Long
'
'    ' Get folder and pidl from path, pidl, or special item
'    Set folder = FolderFromItem(vItem, pidlMenu)
'    If folder Is Nothing Then Exit Function
'
'    ' Get an IContextMenu object
'    On Error GoTo ContextPopMenuFail
'    folder.GetUIObjectOf hwnd, 1, pidlMenu, iidContextMenu, 0, menu
'
'    ' Create an empty popup menu and initialize it with QueryContextMenu
'    hMenu = CreatePopupMenu
'    On Error GoTo ContextPopMenuFail2
'    menu.QueryContextMenu hMenu, 0, 1, &H7FFF, CMF_EXPLORE
'
'    ' Convert x and y to client coordinates
'    ClientToScreenXY hwnd, x, y
'
'    ' Display the context menu
'    Const afMenu = TPM_LEFTALIGN Or TPM_LEFTBUTTON Or _
'                   TPM_RIGHTBUTTON Or TPM_RETURNCMD
'    iCmd = TrackPopupMenu(hMenu, afMenu, x, y, 0, hwnd, ByVal hNull)
'
'    ' If a command was selected from the menu, execute it.
'    If iCmd Then
'        ici.cbSize = LenB(ici)
'        ici.fMask = 0
'        ici.hwnd = hwnd
'        ici.lpVerb = iCmd - 1
'        ici.lpParameters = pNull
'        ici.lpDirectory = pNull
'        ici.nShow = SW_SHOWNORMAL
'        ici.dwHotKey = 0
'        ici.hIcon = hNull
'        menu.InvokeCommand ici
'        ContextPopMenu = True
'    End If
'
'ContextPopMenuFail2:
'    DestroyMenu hMenu
'
'ContextPopMenuFail:
'    ' Menu pidl is freed, so client had better not pass only copy
'    Allocator.Free pidlMenu
'    BugMessage Err.Description
'
'End Function
'DWORD CALLBACK CopyProgressRoutine(
'  __in      LARGE_INTEGER TotalFileSize,
'  __in      LARGE_INTEGER TotalBytesTransferred,
'  __in      LARGE_INTEGER StreamSize,
'  __in      LARGE_INTEGER StreamBytesTransferred,
'  __in      DWORD dwStreamNumber,
'  __in      DWORD dwCallbackReason,
'  __in      HANDLE hSourceFile,
'  __in      HANDLE hDestinationFile,
'  __in_opt  LPVOID lpData
');
'ugh, haven't written callbacks for a while- can't remember wether to use byval or not...
'#

Public Function CopyProgressRoutine(ByVal TotalFileSize As Currency, ByVal TotalBytesTransferred As Currency, _
ByVal StreamSize As Currency, ByVal StreamBytesTransferred As Currency, ByVal dwStreamNumber As Long, _
ByVal dwCallbackReason As Long, ByVal hSourceFile As Long, ByVal hDestinationFile As Long, ByVal lpData As Long) As Long
            
            
          
            
            
            'EDIT: new method:
            'use a module level array- "mCopyProgresscallbacks" or something.
            'have "AddCallback" routine; AddCallback returns a Long that is the index into the array, which is then passed (by the Internal functions that use CopyProgressRoutine) in the data argument.
            'Then "dereferencing" is simply accessing a Array.
            'One drawback: we cannot actually remove any item from the array, we can set the used elements to Nothing... but that's about it.
            
            'step one: convert large integer arguments to doubles.
            Dim dFileSize As Double, dBytesTransferred As Double, dStreamSize As Double, dStreamTransferred As Double
            Dim mCallback As IProgressCallback, SourceFile As CFile, DestFile As CFile
            Dim sSource As String, SDest As String
            Dim gotdata As BCCOPYFILEDATA
            On Error GoTo ErrCopyProgressRoutine
            Debug.Print "copyprogressroutine! lpdata=" & lpData
            'Stop
            'FileSize = LI2DBL(TotalFileSizeHigh, TotalFileSizeLow)
            'BytesTransferred = LI2DBL(TotalBytesTransferredHigh, TotalBytesTransferredLow)
            'StreamSize = LI2DBL(StreamSizeHigh, StreamSizeLow)
            'StreamTransferred = LI2DBL(StreamBytesTransferredHigh, StreamBytesTransferredLo)
            'CDebug.Post "copyprogressroutine"
            dFileSize = TotalFileSize * 10000
            dBytesTransferred = TotalBytesTransferred * 10000
            dStreamSize = StreamSize * 10000
            dStreamTransferred = StreamBytesTransferred * 10000
            'whew.
            Debug.Print "CopyProgressRoutine Filesize=" & dFileSize & ", bytestransferred=" & dBytesTransferred & ", Streamsize=" & _
            dStreamSize & ", StreamTransferred=" & dStreamTransferred
            'dwData will contain address of the object that invoked the filecopy.
            sSource = GetFileNameFromHandle(hSourceFile)
            SDest = GetFileNameFromHandle(hDestinationFile)
            Set SourceFile = FileSystem.GetFile(sSource)
            If SDest <> "" Then
            Set DestFile = FileSystem.GetFile(SDest)
            End If
            'lpdata is index into module level array.
            Set mCallback = mCopyProgresscallbacks(lpData)
            Dim ret As CopyProgressRoutineReturnConstants
            ret = mCallback.UpdateProgress(sSource, SDest, dwCallbackReason, CDbl(dFileSize), CDbl(dBytesTransferred), CDbl(StreamSize), CDbl(dStreamTransferred))
            CopyProgressRoutine = ret
                                    
            Exit Function
ErrCopyProgressRoutine:
            Debug.Print "CopyProgressRoutine error:" & Err.Number & " " & Err.Description
        
End Function
Public Function FileExists(ByVal PathSpec As String, Optional ByVal AcceptDirectories As Boolean = False) As Boolean
    'returns wether a file exists.
'    Dim hFile As Long
'    hFile = CreateFile(PathSpec, GENERIC_DEVICE_QUERY, 0, ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
'    If hFile > 0 Then
'        FileExists = True
'        CloseHandle hFile
'    Else
'
'
'
'        FileExists = False
'    End If
    'use getfileAttributes...
    Dim Attribs As Long
    Dim w32 As WIN32_FIND_DATA, hfind As Long
    Attribs = GetFileAttributes(PathSpec)
    If Attribs > 0 Then
        If ((Attribs And FILE_ATTRIBUTE_DIRECTORY) = 0) Then
            FileExists = True
        ElseIf (Attribs And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY And AcceptDirectories Then
            FileExists = True
        Else
        
            FileExists = False
        End If
    Else
        'doublecheck...
        hfind = FindFirstFile(PathSpec, w32)
        If Replace$(Trim$(w32.cFileName), vbNullChar, "") <> "" Then
            FileExists = True
        Else
            FileExists = False
        End If
        FindClose hfind
        
    End If
    
End Function
Public Function GetTempFileNameAndPathEx() As String
    Dim tpath As String, tfile As String
    tpath = Space$(2048)
    Call GetTempPath(2047, tpath)
    tpath = Left$(tpath, InStr(tpath, vbNullChar) - 1)
    tfile = GetTempFileNameEx
    
    tfile = IIf(Right$(tpath, 1) <> "\", tpath & "\", tpath) & tfile
    GetTempFileNameAndPathEx = tfile
End Function

Public Function GetTempFileNameEx() As String
    'creates a temporary file name based on GUIDs.
    Dim fname As String
    Dim thebytes(1 To 16) As Byte
    Dim pg As Guid, I As Long
    CoCreateGuid pg
    CopyMemory thebytes(1), pg, 16
    'fname = fname & pg.Data1 / 2
    'fname = fname & pg.Data2 / 3
    'fname = fname & pg.Data3 / 4
    For I = 1 To 16
    fname = fname & Chr$((thebytes(I) Mod 24) + 65)
    Next I
    'fname = fname & pg.Data4(
    GetTempFileNameEx = fname & ".TMP"
End Function

'#define BUFSIZE 512
'
Public Function TrimNull(ByVal Strtrim As String) As String
If InStr(Strtrim, vbNullChar) > 0 Then
TrimNull = Mid$(Strtrim, 1, InStrRev(Strtrim, vbNullChar) - 1)
Else
    TrimNull = Strtrim
End If
End Function
Public Function GetFileNameFromHandle(ByVal FileHandle As Long) As String
'BOOL GetFileNameFromHandle(HANDLE hFile)
'{
'  BOOL bSuccess = FALSE;
'  TCHAR pszFilename[MAX_PATH+1];
'  HANDLE hFileMap;
'
'  // Get the file size.
'  DWORD dwFileSizeHi = 0;
'  DWORD dwFileSizeLo = GetFileSize(hFile, &dwFileSizeHi);
'
'  if( dwFileSizeLo == 0 && dwFileSizeHi == 0 )
'  {
'     printf("Cannot map a file with a length of zero.\n");
'     return FALSE;
'  }
'
'  // Create a file mapping object.
'  hFileMap = CreateFileMapping(hFile,
'                    NULL,
'                    PAGE_READONLY,
'                    0,
'                    1,
'                    NULL);
'




    Dim hFileMap As Long, sFileName As String
    Dim dwFileSizeHi As Long
    Dim dwFileSizeLo As Long, pmem As Long
    Dim breturn As Boolean, nbytes As Long
    Dim sTemp As String, drives() As String, I As Long
    Const PAGE_READONLY As Long = &H2
    Const SECTION_MAP_READ As Long = &H4
    Const FILE_MAP_READ As Long = SECTION_MAP_READ
    
    'before we get too carried away, check if we are on vista or later; if so we can use a built-in function to do this task.
    If IsVistaOrLater Then
        GetFileNameFromHandle = GetpathFromhandle(FileHandle)
        Exit Function
    End If
    
    
    
    
trysize:
    dwFileSizeLo = GetFileSize(FileHandle, dwFileSizeHi)
    If dwFileSizeLo = 0 And dwFileSizeHi = 0 Then
        'can't map zero-length file.
        'So... write to it >:)
        breturn = WriteFile(FileHandle, ".", 1, nbytes, ByVal &H0)
        If breturn = 0 Then
            'Epic FAIL.
            GetFileNameFromHandle = ""
        Else
            GoTo trysize
        End If
        
    Else
    
    
    
    
        hFileMap = CreateFileMapping(FileHandle, ByVal 0, PAGE_READONLY, 0, 1, "")
        If hFileMap Then
            pmem = MapViewOfFile(hFileMap, FILE_MAP_READ, 0, 0, 1)
            If pmem Then
                sFileName = Space$(2048)
                Dim sztemp As String
                If GetMappedFileName(GetCurrentProcess, ByVal pmem, sFileName, 2047) Then
                    'translate path with device name to drive letters.
                    sTemp = vbNullChar & Space(2047)
                    If GetLogicalDriveStrings(2057, sTemp) Then
                        drives = Split(sTemp, vbNullChar)
                        For I = 0 To UBound(drives)
                            If Trim$(drives(I)) <> "" Then
                                sTemp = vbNullChar & Space(2047)
                                QueryDosDevice Left$(drives(I), 2), sTemp, 2048
                                sTemp = Left$(sTemp, InStr(sTemp, vbNullChar) - 1)
                                If Len(sTemp) > 0 Then
                                
                                    If StrComp(sTemp, Left$(sFileName, Len(sTemp)), vbTextCompare) = 0 Then
                                        GetFileNameFromHandle = drives(I) & Mid$(sFileName, Len(sTemp) + 2)
                                        Exit For
                                        
                                    End If
                                    
                                End If
                            End If
                        Next I
                    

                    
                    
                    
                    End If
                
                End If
            
            End If
        End If
        
    End If
    If pmem <> 0 Then UnmapViewOfFile pmem
    If hFileMap <> 0 Then CloseHandle hFileMap
End Function




Public Function GetOpenFileString(ByVal Extension As String)
Dim creg As cRegistry, Filemask As String, defvalue As String
'Input: something such as "*.jpg" or "*.txt"
If Left$(Extension, 1) <> "." Then Extension = "." & Extension
'Registry operations will be used to determine the file type-
'for example- let's take, *.EXE

'look at key, ".EXE"
Filemask = Replace$(Extension, "*", "")
'if the default value of that key is also present in HKEY_CLASSES_ROOT, then grab the default value from that key- otherwise return the default value of this key.
Set creg = New cRegistry
defvalue = creg.ValueEx(hhkey_classes_root, Extension, "", RREG_SZ, "")
If defvalue = "" Then

End If
End Function
Function GetFileIcon(ByVal Spath As String, ByVal iconsize As IconSizeConstants) As Long
    Dim finfo As SHFILEINFO
    Dim lIconType As Long
    Dim attruse As FileAttributeConstants
    Dim Flags As Long
    
    If SmallIcons Is Nothing Then
        Set SmallIcons = New cVBALImageList
        
        SmallIcons.Create
    End If
    If LargeIcons Is Nothing Then
        Set LargeIcons = New cVBALImageList
            LargeIcons.Create
            
        End If
    ' be sure that there is the mbNormalIcon too
   
    ' retrieve the item's icon
    Flags = SHGFI_ATTR_SPECIFIED
    If iconsize = icon_large Then
        Flags = Flags + SHGFI_ICON
    ElseIf iconsize = ICON_SMALL Then
        Flags = Flags + SHGFI_SMALLICON + SHGFI_ICON
    ElseIf iconsize = icon_shell Then
        Flags = Flags + SHGFI_SHELLICONSIZE + SHGFI_ICON
    End If

    SHGetFileInfo Spath, attruse, finfo, Len(finfo), Flags
    ' convert the handle to a StdPicture
    GetFileIcon = finfo.hIcon
End Function
Public Function GetPathDepth(ByVal ppath As String) As Long
    'returns the depth of the specified file/folder.
    
    
    Dim Spath As String
    'parse the path...
    
    
    ParsePathParts ppath, , Spath, , , , Parse_Path
    
    'now, count the slashes. remove trailing slash if present.
    
    If Right$(Spath, 1) = "\" Then Spath = Mid$(Spath, 1, Len(Spath) - 1)
    GetPathDepth = Len(Spath) - Len(Replace$(Spath, "\", "")) + 2



End Function
Public Sub unittest()

    Dim str() As String
    ReDim str(0 To 1)
    str(0) = "CDRID.JPG"
    str(1) = "CDRID.gif"
    ShowExplorerMenuMulti FrmDebug.hwnd, "C:\", str()



End Sub
Public Sub TestVolumes()
    Dim currvolume As CVolume
    Dim vols As Volumes
    Set vols = FileSystem.Getvolumes
    Set currvolume = vols.GetNext
    Do
        If currvolume.IsReady Then
        Debug.Print "VOL:" & currvolume.RootFolder.Path
        End If
        
        Set currvolume = vols.GetNext
    
    Loop Until currvolume Is Nothing



End Sub
Public Sub testFiledialog()

    Dim grabfile As CFile
    Dim useopen As CFileDialog
    Set useopen = New CFileDialog
    Set grabfile = useopen.GetFileDirect(FrmDebug.hwnd, , OFN_EXPLORER + OFN_DONTADDTORECENT + OFN_ENABLEHOOK)
    


End Sub
Public Function FindNextFile(ByVal hFile As Long, Win32Data As WIN32_FIND_DATA) As Long
'Static wStruct As WIN32_FIND_DATAW
'Debug.Print "FindNextFile"
If MakeWideCalls Then
    ZeroMemory Win32Data, Len(Win32Data)
    'Debug.Print "Calling:"
    FindNextFile = FindNextFileW(hFile, Win32Data)
    
    'CopyMemory Win32Data, wStruct, Len(Win32Data)
   
    'LSet Win32Data = wStruct
    'BUG: the two buffers run into each other for files with short names...
       Win32Data.cFileName = StrConv(Win32Data.cFileName, vbFromUnicode)
    Win32Data.cFileName = Left$(Win32Data.cFileName, InStr(Win32Data.cFileName, vbNullChar))
    'Debug.Print "FindNextFile Found " & Trim$(Win32Data.cFileName)
Else

    FindNextFile = FindNextFileA(hFile, Win32Data)



End If




End Function

Public Function FindFirstFile(ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
'Dim wStruct As WIN32_FIND_DATAW
'Debug.Print "FindFirstFile..."
'Debug.Print "FindFirstFile"
If MakeWideCalls Then
    'ReDim wStruct.Buffer(1 To 16768 + 1) 'TODO:// special handling via //?/ or whatever that damn prefix is.
    If Not IsUNCPath(lpFileName) Then
        'If InStr(lpFileName, "*") = 0 Then
        lpFileName = "\\?\" & lpFileName
        'End If
    End If
    FindFirstFile = FindFirstFileW(StrPtr(lpFileName), lpFindFileData)
'\\?\
    'lpFindFileData.dwFileAttributes = wStruct.dwFileAttributes
  
    'LSet lpFindFileData = wStruct
    'plop on some null characters, since that will be what the caller expects.
    'lpFindFileData.cFileName = Replace$(wStruct.Buffer, " ", vbNullChar)
    lpFindFileData.cFileName = StrConv(lpFindFileData.cFileName, vbFromUnicode)
    lpFindFileData.cFileName = Left$(lpFindFileData.cFileName, InStr(lpFindFileData.cFileName, vbNullChar))
    
Else
    FindFirstFile = FindFirstFileA(lpFileName, lpFindFileData)

End If




End Function
Public Sub Testalternate(Openme As String)
'D:\outtext.txt
Dim hFile As Long, streams() As String, StrJoin As String, I As Long

Dim testA As CAlternateStreams
Dim testb  As CAlternateStreams
Set testA = GetAlternateStreamsByPath(Openme)
Set testb = ListStreams(Openme)
Stop
'hFile = CreateFile( m_sFile.GetBuffer(0),
'                        GENERIC_READ,
'                        FILE_SHARE_READ,
'                        NULL,
'                        OPEN_EXISTING,
'                        FILE_FLAG_BACKUP_SEMANTICS,
'                        NULL );
'hFile = CreateFile(Openme, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
'streams() = GetAlternateStreamsByPath(hFile)
'For I = 0 To UBound(streams)
'    strjoin = strjoin & vbCrLf & streams(I)
'Next I
'MsgBox "Streams in File """ & Openme & """:" & vbCrLf & strjoin
'CloseHandle hFile
End Sub
Public Function GetAlternateStreamsByPath(ByVal OfFileName As String) As CAlternateStreams

    Dim streamStruct As WIN32_STREAM_ID, pContext As Long
    Dim OfFile As Long
    Dim StreamsRet As CAlternateStreams, newstream As CAlternateStream
    Dim lowpart As Long, highpart As Long, lowseeked As Long, highseeked As Long
    
    Dim dwbytestoread As Long, dwbytesread As Long, lpbytes() As Byte
    Dim bresult As Boolean, badata(4096) As Byte, callresult As Long, SStream() As String, streamCount As Long
    Set StreamsRet = New CAlternateStreams
    
    OfFile = CreateFile(OfFileName, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
    If OfFile = -1 Then
        RaiseAPIError Err.LastDllError, "GetAlternateStreamByPath "
        'Exit Function
    End If
    StreamsRet.Owner = OfFileName
    'dwBytesToRead = Len(streamStruct)
    'ReDim badata(Len(streamStruct) * 2 + 2)
    'bresult = True
    Do
'    if(!BackupRead( hFile,
'                        baData,
'                        dwBytesToRead,
'                        &dwBytesRead,
'                        FALSE, // am I done?
'                        FALSE,
'                        &pContext ) )
'*** "streams" appear to be at the end of files, not the beginning where the header is...
'dwBytesToRead = Len(streamStruct)
'seek to the end of the file- this is the file header+the file length.
'file header is 21 bytes(?)...
    'lowpart = GetFileSize(ofFile, highpart)
    'lowpart = lowpart
    'callresult = BackupSeek(ofFile, lowpart, highpart, lowseeked, highseeked, pContext)
    dwbytestoread = Len(streamStruct)
    dwbytesread = 0
        callresult = BackupRead(OfFile, badata(0), dwbytestoread, dwbytesread, False, False, pContext)
            If Not CBool(callresult) Then
                bresult = False
            Else
                If dwbytesread = 0 Then
                    'all done...
                    bresult = False
                
                Else
                
                
                    bresult = True
                    CopyMemory streamStruct, badata(0), Len(streamStruct)
                    'we read the stream header successfully.
                    '... now we need to Stream name:
                    If LI2DBL(streamStruct.dwStreamSizeLow, streamStruct.dwStreamSizeHigh) > 0 Then
                        'size is not zero- read that data into a string...
                        ReDim Preserve SStream(streamCount)
                        ReDim lpbytes(0 To streamStruct.dwStreamNameSize)
                        'read name...
                        If streamStruct.dwStreamNameSize > 0 Then
                            BackupRead OfFile, lpbytes(0), streamStruct.dwStreamNameSize, dwbytesread, False, False, pContext
                    
                            '****
                            SStream(streamCount) = lpbytes() 'grab the stream name, and store it in the return array.
                            '**NOTES**
                            SStream(streamCount) = Trim$(SStream(streamCount))
                            
                            Set newstream = New CAlternateStream
                            newstream.Init OfFileName, SStream(streamCount), streamStruct.dwStreamSizeLow, streamStruct.dwStreamSizeHigh, streamStruct.dwStreamAttributes
                            'create a collection and Class, "AlternateStreams" and "CAlternateStreams" or something. Here we would initialize the new
                            'CAlternateStreams Object as appropriate; note that "CAlternateStream" should contain methods similar to those of a file-
                            'path would be the entire path to the stream, etc.
                            'It' might even make sense to encapsulate the same functionality into CFile, except that might make it messy....
                            StreamsRet.Add newstream
                            streamCount = streamCount + 1
                        End If
                
                    End If
                    'seek to the end of the stream...
                    BackupSeek OfFile, streamStruct.dwStreamSizeLow, streamStruct.dwStreamSizeHigh, lowseeked, highseeked, pContext
                 End If
                
            End If
        
    Loop While bresult
    'allow backupread to deallocate...
    BackupRead OfFile, 0, 0, 0, 1, 0, pContext
    '    GetAlternateStreamsByHandle = SStream
    Set GetAlternateStreamsByPath = StreamsRet
End Function
Public Sub testpathparts(pathtest As String)

Dim vol As String, pth As String, fname As String, Extension As String, StreamName As String

ParsePathParts pathtest, vol, pth, fname, Extension, StreamName

MsgBox "parts of " & pathtest & vbCrLf & _
    "Volume:" & vol & vbCrLf & _
    "Path:" & pth & vbCrLf & _
    "Filename:" & fname & vbCrLf & _
    "Extension:" & Extension & vbCrLf & _
    "StreamName:" & StreamName

End Sub
'Path Parsing Routines.
Public Function IsUNCPath(pathtest As String) As Boolean
    IsUNCPath = (Left$(pathtest, 2) = "//")
End Function

'Public Function ParseProtocolSpec(ByVal StrInput As String, Optional ByRef Protocol As String, Optional ByRef Site As String, _
'    Optional ByRef Path As String, Optional ByVal Filename As String)
'
''    "parses a protocol string/ URL. For example in:
'
''"http://www.google.ca/search?hl=en&client=firefox-a&channel=s&rls=org.mozilla%3Aen-US%3Aofficial&hs=r0G&q=Protocol+Parsing&btnG=Search&meta="
'
''the protocol would be "HTTP://"
'
''site would be "www.google.ca"
'
''file would be the rest.
'
'
'StrInput = Replace$(StrInput, "\", "/")
'Dim ColonFound As Long
'Dim nextSlash As Long
'
'ColonFound = InStr(StrInput, "://")
'
'Protocol = Mid$(StrInput, 1, ColonFound + 2)
'
''Site will be everything between ColonFound+3 and the Next "/"
'nextSlash = InStr(Len(Protocol) + 1, StrInput, "/", vbTextCompare)
'Site = Mid$(StrInput, ColonFound + 3, nextSlash - (ColonFound + 3))
'
'
'
'
'
'
'
'    End Function
Public Function MakeAbsolutePath(ByVal strpath As String) As String

'An absolute path has the following:
'a Drive specification (or UNC share spec)
'a Path.
'a Filename.

'Step one: determine if we have a volume name. if there is a Colon as the second character
'or two slashes




End Function
Public Function SplitPath(ByVal StrPathName As String) As String()
    'INPUT: a pathname. For example, "C:\filename\testfile\test\file.txt"
    
    'returns: a split array of components of that path.
    
    StrPathName = Replace$(StrPathName, "/", "\")
    
    Do Until InStr(StrPathName, "\\") = 0
        StrPathName = Replace$(StrPathName, "\\", "\")
    Loop
    
    If Left$(StrPathName, 1) = "\" Then StrPathName = Mid$(StrPathName, 2)
    If Right$(StrPathName, 1) = "\" Then StrPathName = Mid$(StrPathName, 1, Len(StrPathName) - 1)


    SplitPath = Split(StrPathName, "\")

End Function
Public Function ParsePathParts(ByVal StrInput As String, Optional ByRef Volume As String, Optional ByRef Path As String, _
Optional ByRef Filename As String, Optional ByRef Extension As String, Optional ByRef StreamName As String, Optional ByVal ParseLevel As ParsePathPartsConstants = Parse_All)
    Dim flpathUNC As Boolean, Countslash As Long, Currpos As Long
    Dim lastSlash As Long, IsProtocol As Boolean
    flpathUNC = IsUNCPath(StrInput)
    StrInput = Replace$(StrInput, "/", "\")
    'Parses a path specification into it's constituent parts.
    
    
    
    
    'Retrieve the volume.
    'Volume could be a drive letter:
    
    'C:\
    
    'or a UNC volume:
    
    '//servername/share/
    
    
    'also, could have a protocol:
    '<protocol>://
    
'    'if we find a Colon followed by two slashes, we will assume that everything before it is the protocol- and everything until the first slash is the "volume"
'    If InStr(1, StrInput, ":\\", vbTextCompare) <> 0 Then
'        'Call my parseProtocol routine...
'    End If
'
'
    
    
    'Note that SHARE is technically part of the drive specification and so should be part of the returned volume name.
    
    'Simple: if it is a UNC path, return the string up to the fourth slash, otherwise the first three characters comprise the volume portion of the path.
    Currpos = 0
    Countslash = 0
    If flpathUNC Then
        Do
            Currpos = InStr(Currpos + 1, StrInput, "\", vbTextCompare)
            Countslash = Countslash + 1
            
        Loop Until Countslash >= 4
        
        Volume = Mid$(StrInput, 1, Currpos)
    
    'ElseIf IsProtocol Then
    'a Protocol-
    
    
    
    
    Else 'neither a protocol or a UNC path.
        'first three...
        Volume = Mid$(StrInput, 1, 3)
        Currpos = 3
    End If
    'ADDED: logic to save time when caller doesn't want certain values.
    If ParseLevel >= Parse_Volume Then
        
        'Path... starts at currpos, lasts until last slash.
        Path = Mid$(StrInput, Len(Volume) + 1, InStrRev(StrInput, "\") - Len(Volume))
        
        
        If ParseLevel >= Parse_Path Then
            
            Dim LastColon As Long
            Dim lastDot As Long
            'Filename and stream
            lastSlash = InStrRev(StrInput, "\")
            lastDot = InStrRev(StrInput, ".")
            If lastDot = 0 Then
                Filename = Mid$(StrInput, lastSlash + 1)
                Extension = ""
                StreamName = ""
            Else
                LastColon = InStrRev(StrInput, ":")
                If LastColon > lastSlash Then
                    Filename = Mid$(StrInput, lastSlash + 1, lastDot - lastSlash - 1)
                    
                    Extension = Mid$(StrInput, lastDot + 1, LastColon - lastDot - 1)
                    
                    StreamName = Mid$(StrInput, LastColon + 1)
                Else
                    Filename = Mid$(StrInput, lastSlash + 1, lastDot - lastSlash - 1)
                    
                    Extension = Mid$(StrInput, lastDot + 1)
                End If
                
            
            
                
            End If
        Else
            'ParseLevel >=ParsePath
        
        End If
    Else
        'Parselevel>=volume...
    End If
    'if we were a UNC path, replace back...
    If flpathUNC Then
     Path = Replace$(Path, "\", "/")
     Volume = Replace$(Volume, "\", "/")
    
    End If

End Function

Public Function SumAscii(ByVal Strsum As String)
    Dim I As Long
    Dim Currsum As Double
    For I = 1 To Len(Strsum)
        Currsum = Currsum + AscW(Mid$(Strsum, I, 1))
        Currsum = Currsum Mod 32768
    Next I
    SumAscii = Currsum



End Function
'FileChangeNotify Thread routine.
Public Function ThreadProcNotify(ByVal lpParameter As Long) As Long

Dim mCopyTo As CFileChangeNotify, ret As Long
'Set mCopyTo = New CFileChangeNotify
'four bytes for object.
CopyMemory mCopyTo, lpParameter, 4
'can't call library functions here. tssk tssk to me.
'just waitforsingleObject our parameter- and set the flag and get the F out of here.
ret = WaitForSingleObject(mCopyTo.ThreadhEvent, &HFFFFFF)


Call ZeroMemory(mCopyTo, 4)

End Function
'DWORD WINAPI ThreadProc(
'  __in  LPVOID lpParameter
');

Public Sub unittest2()
Dim ffile As CFile, fstream As FileStream
Dim ads As CAlternateStream
Set ffile = FileSystem.CreateFile("D:\blob5.bin")
Set ads = ffile.AlternateStreams(False).CreateStream("testit")
    Set fstream = ads.OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_READ, OPEN_ALWAYS, 0)
    fstream.WriteString "THIS IS A STRING", StrRead_ANSI
    fstream.Flush
    fstream.CloseStream
    FileSystem.GetFile("D:\blob3.bin").AlternateStreams(True).count


End Sub
Public Function JoinStr(StrJoin() As String, Optional ByVal Delimiter As String = ",") As String
    Dim builder As cStringBuilder
    Dim currItem As Long
    Set builder = New cStringBuilder
    On Error GoTo JoinFail
    For currItem = LBound(StrJoin) To UBound(StrJoin)
    
        builder.Append StrJoin(currItem)
        If currItem < UBound(StrJoin) Then
            builder.Append Delimiter
        End If
    Next currItem

    JoinStr = builder.ToString
    Set builder = Nothing
    Exit Function
JoinFail:
    JoinStr = ""
End Function
Public Sub testoperation()
    Dim srcfiles() As String
    Dim destfiles() As String
    ReDim srcfiles(1 To 2)
    ReDim destfiles(1 To 1)
    srcfiles(1) = "D:\XPtray.jpg"
    srcfiles(2) = "D:\VistaTray.jpg"
    destfiles(1) = "D:\VBProj\"
    
    FileOperation srcfiles, destfiles, FO_COPY, FOF_ALLOWUNDO + FOF_CONFIRMMOUSE, 0, "progress."

End Sub
Public Function OpenWith(ByVal hWndOwner As String, ByVal FilePath As String, Optional ByVal Icon As Long = 0) As Long

    Dim sei As SHELLEXECUTEINFOA
    Dim verba As String
    verba = "openas"
    sei.hIcon = Icon
    sei.hProcess = GetCurrentProcess()
    sei.hInstApp = App.hInstance
    
    sei.lpVerb = verba
    sei.lpFile = FilePath
    sei.nShow = vbNormalFocus
    sei.cbSize = Len(sei)
    
    OpenWith = ShellExecuteEx(sei)


    'Stop

End Function
Public Function FileOperation(SourceFiles() As String, destfiles() As String, SHop As olelib.FILEOP, Flags As olelib.FILEOP_FLAGS, Optional ByVal OwnerWnd As Long = 0, Optional ByVal ProgressTitle As String = "")

Dim fbuf() As Byte
Dim Foperation As SHFILEOPSTRUCT
Dim srcStr As String, destStr As String
Dim ret As Long
'first, create null-delimited lists for sourcefiles() and destFiles()...

srcStr = JoinStr(SourceFiles(), vbNullChar) & vbNullChar & vbNullChar
destStr = JoinStr(destfiles(), vbNullChar) & vbNullChar & vbNullChar

'alright- we have source and dest...
With Foperation
    .pFrom = srcStr
    .pTo = destStr
    .hwnd = OwnerWnd
    .wFunc = SHop
    .fFlags = Flags
End With

ReDim fbuf(1 To Len(Foperation) + 2)
        ' Now we need to copy the structure into a byte array
    'Call CopyMemory(fbuf(1), Foperation, Len(Foperation))

            ' Next we move the last 12 bytes by 2 to byte align the data
    'Call CopyMemory(fbuf(19), fbuf(21), 12)
    'last- call routine...
    
    
    ret = SHFileOperation(Foperation)
    'copy last 12 bytes back...
    'CopyMemory fbuf(21), fbuf(19), 12
    'copy into structure...
    'CopyMemory Foperation, fbuf(1), UBound(fbuf)
    Stop
    If ret <> 0 Then
        RaiseAPIError Err.LastDllError, "MdlFileSystem::FileOperation"
        
    Else
        '
    
    End If
'            If result <> 0 Then  ' Operation failed
'               MsgBox Err.LastDllError 'Show the error returned from
'                                       'the API.
'               Else
'               If FILEOP.fAnyOperationsAborted <> 0 Then
'                  MsgBox "Operation Failed"
'               End If
'            End If




End Function


'more utility functions- mostly dealing with Paths and Shell path handling routines.

Public Function MakePathAbsolute(ByVal RelativePath As String, Optional ByVal RelativeTo As String = "") As String
'
    Dim isunc As Boolean
   If IsUNCPath(RelativePath) Then isunc = True
    Dim BuildArray() As String
    Dim arraySize As Long
    Dim returnString As String, getvol As String
   Dim splPathParts() As String
   Dim I As Long
   
   If Mid$(RelativePath, 2, 1) = ":" Or _
    Left$(RelativePath, 2) = "//" Then
    'it's a volume name, so return the relative string unabated.
    getvol = MdlFileSystem.GetVolume(RelativePath)
    MakePathAbsolute = MakePathAbsolute(Mid$(RelativePath, Len(getvol)), getvol)
    Exit Function
  End If
   
   
    RelativePath = FixPath(RelativePath)
   ' RelativeTo = Replace$(RelativeTo, "/", "\")
   'so- the inevitable question is what needs to be done?
   'First, parse the path parts of both given paths. the RelativePath might have a filename/stream info, so we need to preserve that.
   
   'first, split our relativepath into it's specific path portions:
   ' relative = Split(RelativePath, "\")
   splPathParts = Split(RelativePath, "\")
   
   
   'we now have two arrays.
   Dim numseps As Long
   
   BuildArray = Split(RelativeTo, "\")
   arraySize = UBound(BuildArray)
   If BuildArray(arraySize) = "" Then
   arraySize = arraySize - 1
    ReDim Preserve BuildArray(arraySize)
    
    End If
   For I = 0 To UBound(splPathParts)
    If StrComp(splPathParts(I), "..") = 0 Then
        'remove the last item from Buildarray...
        If arraySize > 1 Then
            arraySize = arraySize - 1
            ReDim Preserve BuildArray(arraySize)
        End If
    ElseIf StrComp(splPathParts(I), ".") = 0 Then
        'no change...
    ElseIf splPathParts(I) = "" Then
    Else
        'add this to Buildarray...
        arraySize = arraySize + 1
        ReDim Preserve BuildArray(arraySize)
        BuildArray(arraySize) = splPathParts(I)
    
    End If
   
   
   
   
   Next I
   
   
    'now- reiterate and rebuild a path from buildarray().
    For I = 0 To UBound(BuildArray)
        returnString = returnString & BuildArray(I)
        If I < UBound(BuildArray) Then
        returnString = returnString & "\"
        End If
    
    Next I
    MakePathAbsolute = returnString


End Function

Public Sub CopyEntireStream(SourceStream As IInputStream, destStream As IOutputStream, Optional ByVal Chunksize As Long = 32768)
    
    If SourceStream.Size < Chunksize Then
    Debug.Print "setting chunk size to " & SourceStream.Size
        Chunksize = SourceStream.Size
    End If
    SourceStream.SeekTo 0, STREAM_BEGIN
    Do Until SourceStream.EOF
        destStream.WriteBytes SourceStream.readbytes(Chunksize)

    '
    Loop

End Sub
Public Sub testunit()
'Unit test.

    Dim persistit As Object
    Dim fstream As FileStream
    Set persistit = CreateObject("Test.Persistor")
    
    Set fstream = FileSystem.CreateFile("D:\persistence2.dat").OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_WRITE, OPEN_EXISTING)
    fstream.WriteObject persistit
    
    fstream.CloseStream
    Set persistit = Nothing
    
    Set fstream = FileSystem.GetFile("D:\persistence2.dat").OpenAsBinaryStream(GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING)
    Set persistit = fstream.ReadObject



End Sub

'Public Sub pleasenocrash()



'End Sub
Private Function Modulus(ByVal Dividend As Variant, ByVal Divisor) As Variant
    'returns the modulus of the two numbers.
    'The modulus is the remainder after dividing Divisor and dividend.
    Dim Quotient As Variant
    'It will, of course, be less then divisor.
    'IE:
    '10 mod 3.3 should be
    Quotient = CDec(Dividend / Divisor)
    'the floating point portion will be
    'the percentage of the divisor that fit at the end.
    Modulus = CDec(Quotient - Int(Quotient)) * CDec(Divisor)
    




End Function



Public Sub testcopy()
    Dim testin As FileStream, testout As FileStream
    
    Set testin = FileSystem.GetFile("D:\nsmrec3_converted.mp4").OpenAsBinaryStream(GENERIC_READ, FILE_SHARE_WRITE, OPEN_EXISTING)
    Set testout = FileSystem.CreateFile("D:\nsm2.mp4").OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_WRITE, OPEN_EXISTING)

    StreamCopy testin, testout, Nothing
    testout.CloseStream
End Sub
Public Function StreamCopy(inputstream As IInputStream, outputstream As IOutputStream, Optional Callback As ICopyMoveCallback) As Long

    Dim Chunksize As Long
    Dim CurrChunk As Long, TotalChunks As Long, ChunKRemainder As Long
    Dim inputsize As Double
    Dim thisChunkSize As Double
    Dim BytesTransfer() As Byte
    Chunksize = 32& * 1024&
    If Not inputstream.Valid Then
        Err.Raise 5, "MdlFileSystem::StreamCopy", "Passed InputStream Not Valid."
    ElseIf outputstream.Valid Then
        Err.Raise 5, "MdlFileSystem::StreamCopy", "Passed OutputStream not valid."
    
    End If
    If Callback Is Nothing Then Set Callback = New ICopyMoveCallback
    Callback.InitCopy inputstream, outputstream, Chunksize
    
    'Alright- calculate the number of chunks and the remainder chunk...
    inputsize = inputstream.Size
    TotalChunks = Fix(inputsize / Chunksize)
    
    ChunKRemainder = Round(Modulus(inputsize, Chunksize), 0)
    
    'Loop from currchunk to TotalChunks+1....
    thisChunkSize = Chunksize
    For CurrChunk = 1 To TotalChunks + 1
        If CurrChunk > TotalChunks Then thisChunkSize = ChunKRemainder
        BytesTransfer = inputstream.readbytes(thisChunkSize)
        outputstream.WriteBytes BytesTransfer
        Callback.StreamProgress inputstream, outputstream, Chunksize, CurrChunk, CLng(inputsize)
        
        
    
    
        
    
    Next CurrChunk
    
    
    




End Function




Public Sub TestNetwork()
Call FileSystem.CreateFile("//Satellite/Shared/File.txt").OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_READ, OPEN_EXISTING).WriteString(Space$(65536))

End Sub
Public Sub TestXor()
    Dim UseInput As CMemStream, useoutput As CMemStream
    Dim usestring As String
    Dim usefilter As CCoreFilters, FilterObj As IStreamFilter
    Set usefilter = New CCoreFilters
    usefilter.FilterType = Encrypt_CryptoAPI
    usefilter.password = "chopper"
    Set FilterObj = usefilter
    Set UseInput = FileSystem.CreateMemoryStream
    Set useoutput = FileSystem.CreateMemoryStream
    
    
    usestring = InputBox$("Enter string to en/decode:")
    UseInput.WriteString usestring, StrRead_ANSI
    UseInput.SeekTo CDbl(0), STREAM_BEGIN
    FilterObj.FilterStream UseInput, useoutput
    Call useoutput.SeekTo(CDbl(0), STREAM_BEGIN)
    Dim outputstring As String, Result As String
    outputstring = useoutput.ReadAllStr(StrRead_ANSI)
    Debug.Print "encoded string is:" & outputstring
    
    
    useoutput.SeekTo CDbl(0), STREAM_BEGIN
    Set UseInput = FileSystem.CreateMemoryStream()
    usefilter.password = "chopper"
    usefilter.FilterType = Decrypt_CryptoAPI
    FilterObj.FilterStream useoutput, UseInput
    
    UseInput.SeekTo CDbl(0), STREAM_BEGIN
    
    Result = UseInput.ReadAllStr
    
    Debug.Print "decoded is " & Result
    
    
    UseInput.CloseStream
    useoutput.CloseStream
    Set UseInput = Nothing
    Set useoutput = Nothing
    




End Sub
Public Sub TestXorEncode()
    Dim usestream As FileStream, Outstream As FileStream
    Dim usefilter  As CCoreFilters
    Dim Filter As IStreamFilter
    
    Set usefilter = New CCoreFilters
    usefilter.FilterType = encodedecode_xor
    usefilter.password = "chopper"
    Set Filter = usefilter
    Set usestream = FileSystem.OpenStream("D:\hijack.txt")
    Set Outstream = FileSystem.CreateStream("D:\hijack.xor")
    Filter.FilterStream usestream, Outstream
    Outstream.CloseStream
    usestream.CloseStream
    Set usestream = FileSystem.OpenStream("D:\hijack.xor")
    Set Outstream = FileSystem.CreateStream("D:\hijack-2.txt")
    Filter.FilterStream usestream, Outstream
    usestream.CloseStream
    Outstream.CloseStream
End Sub
Public Sub testcompressor()
Dim lzwin As FileStream, lzwout As FileStream
Dim Filter As IStreamFilter, newfilt As CCoreFilters
Set lzwin = FileSystem.OpenStream("D:\zwicky.txt")
Set lzwout = FileSystem.CreateStream("D:\crypted.txt")


Set newfilt = New CCoreFilters
newfilt.FilterType = Encrypt_CryptoAPI
Set Filter = newfilt
Filter.FilterStream lzwin, lzwout

lzwin.CloseStream
lzwout.CloseStream
Set newfilt = New CCoreFilters
newfilt.FilterType = Decrypt_CryptoAPI
Set Filter = newfilt

Set lzwin = FileSystem.OpenStream("D:\crypted.txt")
Set lzwout = FileSystem.CreateStream("D:\decr_exp2.txt")
Filter.FilterStream lzwin, lzwout
lzwin.CloseStream
lzwout.CloseStream
End Sub



Public Sub test()
Dim sects() As String, Scount As Long
'    mreg.ClassKey = HHKEY_USERS
'    mreg.Machine = "SATELLITE"
'    mreg.SectionKey = ""
'    mreg.EnumerateSections sects(), scount
'    Stop
'  mreg.ClassKey = HHKEY_USERS
    mreg.Machine = "SATELLITE"
'    mreg.SectionKey = ".DEFAULT\Control Panel\Colors"
'    'mreg.EnumerateSections sects(), scount
'    mreg.EnumerateValues sects(), scount
'    Stop
mreg.Classkey = HHKEY_USERS
mreg.SectionKey = ".DEFAULT\marksie"
mreg.ValueKey = "ARCIE"
mreg.CreateKey
End Sub
Public Function GetCountStr(ByVal StrCountIn As String, ByVal FindStr As String, Optional Comparemode As VbCompareMethod = vbBinaryCompare) As Long

    Dim Currpos As Long, countof As Long
    Currpos = 1
    countof = 1
    Do
        Currpos = InStr(Currpos + 1, StrCountIn, FindStr, Comparemode)
        If Currpos > 0 Then
            countof = countof + 1
        Else
            If countof = 1 Then
                GetCountStr = 0
                Exit Function
            End If
        End If
    
    
    Loop Until Currpos = 0
    
    
    
GetCountStr = countof



End Function
Public Function GetReducedPath(ByVal strpath As String, ByVal TargetLength As Long) As String





'
'Dim parsedFileName As String, ParsedExtension As String, ParsedStream As String
'Dim ParsedPath As String
'Dim parsedvolume As String
'Dim filename As String
'Dim workString As String
'Dim PrevPath As String
'Dim SplitPath() As String
If Len(strpath) <= TargetLength Then
    GetReducedPath = strpath
    Exit Function
Else
    GetReducedPath = Left$(strpath, TargetLength \ 2) & "..." & Right$(strpath, TargetLength \ 2)


End If




'Dim I As Long, Numiterate As Long
'ParsePathParts StrPath, parsedvolume, ParsedPath, parsedFileName, ParsedExtension, ParsedStream, Parse_All
'If Right$(ParsedPath, 1) = "\" Then ParsedPath = Mid$(ParsedPath, 1, Len(ParsedPath) - 1)
'SplitPath = Split(ParsedPath, "\")
'If ParsedExtension <> "" Or parsedFileName <> "" Then
'    filename = parsedFileName & "." & ParsedExtension
'End If
'workString = parsedvolume
'Dim strStart As String, StrEnd As String
'strStart = parsedvolume
'StrEnd = SplitPath(UBound(SplitPath)) & "\" & filename
'
'For I = 0 To UBound(SplitPath)
'    workString = workString & SplitPath(I) & "\"
'
'Next I




'GetReducedPath = PrevPath

End Function

Public Sub TestPackageHeir()
    Dim testpackage As CHeirarchalFilePackage, createdstream As FileStream
    Set testpackage = New CHeirarchalFilePackage
    testpackage.AddFile "D:\RHDSetup.log", "RHDSetup.log", "\"
    testpackage.AddFile "D:\vcredist.bmp", "vcredist.bmp", "\"
    testpackage.AddFile "H:\test.jpg", "testjpg", "\JPG"
    Set createdstream = FileSystem.CreateStream("C:\outputtestingxxx2.dat")
    testpackage.WriteToStream createdstream
    createdstream.CloseStream
    Set testpackage = New CHeirarchalFilePackage
    testpackage.ReadFromStream FileSystem.OpenStream("C:\outputtestingxxx2.dat")
    Stop
End Sub
'Public Sub TestZip()
'    Dim unzipper As CGUnzipFiles
'    Set unzipper = New CGUnzipFiles
'    unzipper.ExtractList = ListContents
'    unzipper.ExtractDir = "D:\"
'    unzipper.ZipFileName = "D:\vbproj\FreeImage3110.zip"
'    unzipper.Unzip
'    Debug.Print unzipper.GetLastMessage
'
'
'
'
'End Sub
Public Function GetAttributeString(ByVal AttributeValue As FileAttributeConstants, Optional ByVal Longform As Boolean = False) As String
    'returns a comma separated list of the attributes of the file.
    'IE: R,H,S,T etc.
    Static Lookup() As Long, flInit As Boolean
    Static AttrStr() As Variant, attrlook() As Variant
    Static AttrLongStr() As Variant
'    FILE_ATTRIBUTE_ARCHIVE = &H20
'    FILE_ATTRIBUTE_COMPRESSED = &H800
'    FILE_ATTRIBUTE_DEVICE = &H40
'    FILE_ATTRIBUTE_DIRECTORY = &H10
'    FILE_ATTRIBUTE_ENCRYPTED = &H4000
'    FILE_ATTRIBUTE_HIDDEN = &H2
'    FILE_ATTRIBUTE_NORMAL = &H80
'    FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
'    FILE_ATTRIBUTE_OFFLINE = &H1000
'    FILE_ATTRIBUTE_READONLY = &H1
'    FILE_ATTRIBUTE_REPARSE_POINT = &H400
'    FILE_ATTRIBUTE_SPARSE_FILE = &H200
'    FILE_ATTRIBUTE_SYSTEM = &H4
'    FILE_ATTRIBUTE_TEMPORARY = &H100
    Dim attrib As FileAttributeConstants, I As Long, usearray() As Variant
    Dim buildstr As String
    If Not flInit Then
        flInit = True
        ReDim attrlook(1 To 14)
        ReDim Lookup(1 To 14)
        ReDim AttrLongStr(1 To 14)
        Lookup(1) = FILE_ATTRIBUTE_ARCHIVE
        attrlook(1) = "A"
        AttrLongStr(1) = "Archive"
        Lookup(2) = FILE_ATTRIBUTE_COMPRESSED
        attrlook(2) = "C"
        AttrLongStr(2) = "Compressed"
        Lookup(3) = FILE_ATTRIBUTE_DEVICE
        Lookup(4) = FILE_ATTRIBUTE_DIRECTORY
        Lookup(5) = FILE_ATTRIBUTE_ENCRYPTED
        attrlook(5) = "E"
        AttrLongStr(5) = "Encrypted"
        Lookup(6) = FILE_ATTRIBUTE_HIDDEN
        attrlook(6) = "H"
        AttrLongStr(6) = "Hidden"
        Lookup(7) = FILE_ATTRIBUTE_NORMAL
        attrlook(7) = "N"
        AttrLongStr(7) = "Normal"
        Lookup(8) = FILE_ATTRIBUTE_NOT_CONTENT_INDEXED
        Lookup(9) = FILE_ATTRIBUTE_OFFLINE
        Lookup(10) = FILE_ATTRIBUTE_READONLY
        attrlook(10) = "R"
        AttrLongStr(10) = "Read-Only"
        Lookup(11) = FILE_ATTRIBUTE_REPARSE_POINT
        Lookup(12) = FILE_ATTRIBUTE_SPARSE_FILE
        Lookup(13) = FILE_ATTRIBUTE_SYSTEM
        attrlook(13) = "S"
        AttrLongStr(13) = "System"
        Lookup(14) = FILE_ATTRIBUTE_TEMPORARY
    End If
    If Longform Then
        usearray = AttrLongStr
    Else
        usearray = attrlook
    End If
    For I = 1 To 14
        If (AttributeValue And Lookup(I)) = Lookup(I) Then
        If usearray(I) <> "" Then
            buildstr = buildstr & usearray(I) & ","
        End If
        End If
    
    Next I
    If Right$(buildstr, 1) = "," Then buildstr = Left$(buildstr, Len(buildstr) - 1)
    GetAttributeString = buildstr
End Function
Public Sub testallattributes()
    Dim retattrib() As FileAttributeConstants
    Dim I As Long
    retattrib = GetAllFileAttributes()
    
    'Stop
    For I = LBound(retattrib) To UBound(retattrib)
        Debug.Print MdlFileSystem.GetAttributeString(retattrib(I), True)
    Next
    
End Sub
Public Function GetAllFileAttributes() As FileAttributeConstants()

    Dim retArray() As FileAttributeConstants
    
    
    Dim allattributes As Variant
    
    'allattributes = Array(FileAttributeConstants.FILE_ATTRIBUTE_ARCHIVE, FILE_ATTRIBUTE_COMPRESSED, FILE_ATTRIBUTE_DEVICE, FILE_ATTRIBUTE_DIRECTORY, _
    FILE_ATTRIBUTE_ENCRYPTED, FILE_ATTRIBUTE_HIDDEN, FILE_ATTRIBUTE_NORMAL, FILE_ATTRIBUTE_NOT_CONTENT_INDEXED, FILE_ATTRIBUTE_OFFLINE, FILE_ATTRIBUTE_READONLY, _
    FILE_ATTRIBUTE_REPARSE_POINT, FILE_ATTRIBUTE_SPARSE_FILE, FILE_ATTRIBUTE_SYSTEM, FILE_ATTRIBUTE_TEMPORARY)
    allattributes = Array(FILE_ATTRIBUTE_READONLY, FILE_ATTRIBUTE_HIDDEN, FILE_ATTRIBUTE_SYSTEM, FILE_ATTRIBUTE_ARCHIVE, FILE_ATTRIBUTE_NORMAL, FILE_ATTRIBUTE_COMPRESSED, FILE_ATTRIBUTE_ENCRYPTED)
    
    Dim BaseAttribute As FileAttributeConstants
    Dim currBaseIndex As Long
    Dim currloop As Long, CountRunner As Long
    Dim I As Long
    ReDim retArray(1 To UBound(allattributes) + 1)
    For I = 0 To UBound(allattributes)
        retArray(I + 1) = allattributes(I)
    Next I
    'currBaseIndex = UBound(allattributes)
    CountRunner = UBound(allattributes)
    
    Do
        BaseAttribute = BaseAttribute + allattributes(currBaseIndex)
        For currloop = currBaseIndex + 1 To UBound(allattributes)
            CountRunner = CountRunner + 1
            ReDim Preserve retArray(1 To CountRunner)
            retArray(CountRunner) = BaseAttribute + allattributes(currloop)
        Next currloop
    
    
    
        currBaseIndex = currBaseIndex + 1
    Loop Until currBaseIndex = UBound(allattributes)




    GetAllFileAttributes = retArray
End Function


Public Sub TestFilterstack()

'first, create the test file.
Const testfilename = "D:\A.txt"
Const outfilename = "D:\A.dat"
Const reversefilename = "D:\A_reversed.txt"
Const testoutput As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!@#$%^&*()"
Dim writtenstream As FileStream
If FileSystem.Exists(testfilename) Then FileSystem.GetFile(testfilename).Delete
'read a read/write stream...
Set writtenstream = FileSystem.CreateFile(testfilename).OpenAsBinaryStream(GENERIC_READ + GENERIC_WRITE, FILE_SHARE_READ, OPEN_EXISTING)

writtenstream.WriteString testoutput, StrRead_ANSI
writtenstream.WriteString LCase$(testoutput), StrRead_ANSI
'seek to the start.
writtenstream.SeekTo 0


'create an output file...

Dim outstr As FileStream
Dim fstack As CFilterStack
Dim newfilter As CCoreFilters
Set outstr = FileSystem.CreateStream(outfilename)

'now, we create the filter stack...
Set fstack = New CFilterStack
Set newfilter = New CCoreFilters
newfilter.FilterType = Encrypt_CryptoAPI
newfilter.password = "First"
fstack.Add newfilter, "Firstencrypt"

Set newfilter = New CCoreFilters
newfilter.FilterType = Huffman_Compress
newfilter.password = "First"
fstack.Add newfilter


fstack.FilterStream writtenstream, outstr, False

writtenstream.CloseStream
outstr.CloseStream
Dim reversedstream As FileStream
Set outstr = FileSystem.OpenStream(outfilename)
If FileSystem.Exists(reversefilename) Then FileSystem.GetFile(reversefilename).Delete
Set reversedstream = FileSystem.CreateFile(reversefilename).OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_READ, TRUNCATE_EXISTING)

fstack.FilterStream outstr, reversedstream, True

outstr.CloseStream
reversedstream.CloseStream


End Sub
Public Function DOSDateTimetoOLETime(DOSTime As Long, DosDate As Long) As Date
'1 - 4     5 - 10     11 - 16
'day           month      year from 1980

Dim ftime As FILETIME, systime As SYSTEMTIME
DosDateTimeToFileTime DosDate, DOSTime, ftime

FileTimeToSystemTime ftime, systime
DOSDateTimetoOLETime = DateSerial(systime.wYear, systime.wMonth, systime.wDay) + TimeSerial(systime.wHour, systime.wMinute, systime.wSecond)





End Function
