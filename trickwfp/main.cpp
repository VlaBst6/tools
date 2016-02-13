
// http://www.ntcore.com/files/wfp.htm
// released on www.rootkit.com - november the 9th, 2004.
// Daniel Pistelli

// just tested on win2k sp0 to update usbser.sys to the XP sp3 version so I could use
// Arduino on an old win2k machine and it worked perfect :) -dz
// note: permanent patch doesnt work on XP SP3, but live disable does until reboot..

#include <windows.h>
#include <stdio.h>
#include <conio.h>

#ifndef UNICODE_STRING
typedef struct _UNICODE_STRING
{
    USHORT Length;
    USHORT MaximumLength;
    PWSTR  Buffer;
} UNICODE_STRING;
typedef UNICODE_STRING *PUNICODE_STRING;
#endif

#ifndef NTSTATUS
typedef LONG NTSTATUS;
#define NT_SUCCESS(Status) ((NTSTATUS)(Status) >= 0)
#define STATUS_SUCCESS      ((NTSTATUS)0x00000000L)
#endif

#ifndef SYSTEM_INFORMATION_CLASS
typedef enum _SYSTEM_INFORMATION_CLASS {
   SystemBasicInformation,                // 0
   SystemProcessorInformation,            // 1
   SystemPerformanceInformation,          // 2
   SystemTimeOfDayInformation,            // 3
   SystemNotImplemented1,                 // 4
   SystemProcessesAndThreadsInformation,  // 5
   SystemCallCounts,                      // 6
   SystemConfigurationInformation,        // 7
   SystemProcessorTimes,                  // 8
   SystemGlobalFlag,                      // 9
   SystemNotImplemented2,                 // 10
   SystemModuleInformation,               // 11
   SystemLockInformation,                 // 12
   SystemNotImplemented3,                 // 13
   SystemNotImplemented4,                 // 14
   SystemNotImplemented5,                 // 15
   SystemHandleInformation,               // 16
   SystemObjectInformation,               // 17
   SystemPagefileInformation,             // 18
   SystemInstructionEmulationCounts,      // 19
   SystemInvalidInfoClass1,               // 20
   SystemCacheInformation,                // 21
   SystemPoolTagInformation,              // 22
   SystemProcessorStatistics,             // 23
   SystemDpcInformation,                  // 24
   SystemNotImplemented6,                 // 25
   SystemLoadImage,                       // 26
   SystemUnloadImage,                     // 27
   SystemTimeAdjustment,                  // 28
   SystemNotImplemented7,                 // 29
   SystemNotImplemented8,                 // 30
   SystemNotImplemented9,                 // 31
   SystemCrashDumpInformation,            // 32
   SystemExceptionInformation,            // 33
   SystemCrashDumpStateInformation,       // 34
   SystemKernelDebuggerInformation,       // 35
   SystemContextSwitchInformation,        // 36
   SystemRegistryQuotaInformation,        // 37
   SystemLoadAndCallImage,                // 38
   SystemPrioritySeparation,              // 39
   SystemNotImplemented10,                // 40
   SystemNotImplemented11,                // 41
   SystemInvalidInfoClass2,               // 42
   SystemInvalidInfoClass3,               // 43
   SystemTimeZoneInformation,             // 44
   SystemLookasideInformation,            // 45
   SystemSetTimeSlipEvent,                // 46
   SystemCreateSession,                   // 47
   SystemDeleteSession,                   // 48
   SystemInvalidInfoClass4,               // 49
   SystemRangeStartInformation,           // 50
   SystemVerifierInformation,             // 51
   SystemAddVerifier,                     // 52
   SystemSessionProcessesInformation      // 53
} SYSTEM_INFORMATION_CLASS;
#endif

#ifndef HANDLEINFO
typedef struct HandleInfo{
        ULONG Pid;
        USHORT  ObjectType;
        USHORT  HandleValue;
        PVOID ObjectPointer;
        ULONG AccessMask;
} HANDLEINFO, *PHANDLEINFO;
#endif

#ifndef SYSTEMHANDLEINFO
typedef struct SystemHandleInfo {
        ULONG nHandleEntries;
        HANDLEINFO HandleInfo[1];
} SYSTEMHANDLEINFO, *PSYSTEMHANDLEINFO;
#endif

NTSTATUS (NTAPI *pNtQuerySystemInformation)(
  SYSTEM_INFORMATION_CLASS SystemInformationClass,
  PVOID SystemInformation,
  ULONG SystemInformationLength,
  PULONG ReturnLength
);

#ifndef STATUS_INFO_LENGTH_MISMATCH
#define STATUS_INFO_LENGTH_MISMATCH   ((NTSTATUS)0xC0000004L)
#endif

#ifndef OBJECT_INFORMATION_CLASS
typedef enum _OBJECT_INFORMATION_CLASS {
    ObjectBasicInformation,
    ObjectNameInformation,
    ObjectTypeInformation,
    ObjectAllTypesInformation,
    ObjectHandleInformation
} OBJECT_INFORMATION_CLASS;
#endif

#ifndef OBJECT_NAME_INFORMATION
typedef struct _OBJECT_NAME_INFORMATION
{
  UNICODE_STRING ObjectName;

} OBJECT_NAME_INFORMATION, *POBJECT_NAME_INFORMATION;
#endif

#ifndef OBJECT_BASIC_INFORMATION
typedef struct _OBJECT_BASIC_INFORMATION
{
  ULONG                   Unknown1;
  ACCESS_MASK             DesiredAccess;
  ULONG                   HandleCount;
  ULONG                   ReferenceCount;
  ULONG                   PagedPoolQuota;
  ULONG                   NonPagedPoolQuota;
  BYTE                    Unknown2[32];
} OBJECT_BASIC_INFORMATION, *POBJECT_BASIC_INFORMATION;
#endif



NTSTATUS (NTAPI *pNtQueryObject)(IN HANDLE ObjectHandle,
                         IN OBJECT_INFORMATION_CLASS ObjectInformationClass,
                         OUT PVOID ObjectInformation,
                         IN ULONG ObjectInformationLength,
                         OUT PULONG ReturnLength OPTIONAL);


BOOL (WINAPI *pEnumProcesses)(DWORD *lpidProcess, DWORD cb,
                       DWORD *cbNeeded);

BOOL (WINAPI *pEnumProcessModules)(HANDLE hProcess,
                           HMODULE *lphModule,
                           DWORD cb, LPDWORD lpcbNeeded);

DWORD (WINAPI *pGetModuleFileNameExW)(HANDLE hProcess, HMODULE hModule,
                             LPWSTR lpFilename, DWORD nSize);

VOID GetFileName(WCHAR *Name)
{
   WCHAR *path, *New, *ptr;

   path = (PWCHAR) malloc((MAX_PATH + 1) * sizeof (WCHAR));
   New = (PWCHAR) malloc((MAX_PATH + 1) * sizeof (WCHAR));

   wcsncpy(path, Name, MAX_PATH);

   if (wcsncmp(path, L"\\SystemRoot", 11) == 0)
   {
      ptr = &path[11]; 
      GetWindowsDirectoryW(New, MAX_PATH * sizeof (WCHAR));
      wcscat(New, ptr);
      wcscpy(Name, New);
   }
   else if (wcsncmp(path, L"\\??\\", 4) == 0)
   {
      ptr = &path[4];
      wcscpy(New, ptr);
      wcscpy(Name, New);
   }

   free(path);
   free(New);
}

BOOL SetPrivileges(VOID)
{
   HANDLE hProc;
   LUID luid;
   TOKEN_PRIVILEGES tp;
   HANDLE hToken;
   TOKEN_PRIVILEGES oldtp;
   DWORD dwSize;

   hProc = GetCurrentProcess();

   if (!OpenProcessToken(hProc, TOKEN_QUERY |
      TOKEN_ADJUST_PRIVILEGES, &hToken))
      return FALSE;

   if (!LookupPrivilegeValue(NULL, SE_DEBUG_NAME, &luid))
   {
      CloseHandle (hToken);
      return FALSE;
   }

   ZeroMemory (&tp, sizeof (tp));

   tp.PrivilegeCount = 1;
   tp.Privileges[0].Luid = luid;
   tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;

   if (!AdjustTokenPrivileges(hToken, FALSE, &tp, sizeof(TOKEN_PRIVILEGES),
      &oldtp, &dwSize))
   {
      CloseHandle(hToken);
      return FALSE;
   }

   return TRUE;
}

BOOL CompareStringBackwards(WCHAR *Str1, WCHAR *Str2)
{
   INT Len1 = wcslen(Str1), Len2 = wcslen(Str2);


   if (Len2 > Len1)
      return FALSE;

   for (Len2--, Len1--; Len2 >= 0; Len2--, Len1--)
   {
      if (Str1[Len1] != Str2[Len2])
         return FALSE;
   }

   return TRUE;
}

BOOL TrickWFP(bool patchDllToo)
{
   HINSTANCE hNtDll, hPsApi;
   PSYSTEMHANDLEINFO pSystemHandleInfo = NULL;
   ULONG uSize = 0x1000, i, uBuff;
   NTSTATUS nt;

   // psapi variables
   
   LPDWORD lpdwPIDs = NULL;
   DWORD WinLogonId;
   DWORD dwSize;
   DWORD dwSize2;
   DWORD dwIndex;
   HMODULE hMod;
   HANDLE hProcess, hWinLogon;
   DWORD dwLIndex = 0;

   WCHAR Buffer[MAX_PATH + 1];
   WCHAR Buffer2[MAX_PATH + 1];
   WCHAR WinLogon[MAX_PATH + 1];

   HANDLE hCopy;

   // OBJECT_BASIC_INFORMATION ObjInfo; // inutilizzato
   struct { UNICODE_STRING Name; WCHAR Buffer[MAX_PATH + 1]; } ObjName;
   
   OSVERSIONINFOEX osvi;

   HANDLE hFile;
   DWORD dwFileSize, BRW = 0, dwCount;
   BYTE *pSfc, *pCode;
   BOOL bFound = FALSE;

   PIMAGE_DOS_HEADER ImgDosHeader;
   PIMAGE_NT_HEADERS ImgNtHeaders;
   PIMAGE_SECTION_HEADER ImgSectionHeader;

   HKEY Key;
   LONG Ret;

   ZeroMemory(&osvi, sizeof(OSVERSIONINFOEX));
   
   osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
   
   if (!GetVersionEx((OSVERSIONINFO *) &osvi))
   {
      osvi.dwOSVersionInfoSize = sizeof (OSVERSIONINFO);

      if (!GetVersionEx ((OSVERSIONINFO *) &osvi))
         return FALSE;
   }


   if (osvi.dwPlatformId != VER_PLATFORM_WIN32_NT ||
      osvi.dwMajorVersion <= 4)
      return FALSE;

   hNtDll = LoadLibrary("ntdll.dll");
   hPsApi = LoadLibrary("psapi.dll");

   if (!hNtDll || !hPsApi)
      return FALSE;

   // ntdll functions

   pNtQuerySystemInformation = (NTSTATUS (NTAPI *)(
      SYSTEM_INFORMATION_CLASS, PVOID, ULONG, PULONG))
      GetProcAddress(hNtDll, "NtQuerySystemInformation");

   pNtQueryObject = (NTSTATUS (NTAPI *)(HANDLE,
      OBJECT_INFORMATION_CLASS, PVOID, ULONG, PULONG))
      GetProcAddress(hNtDll, "NtQueryObject");

   // psapi functions

   pEnumProcesses = (BOOL (WINAPI *)(DWORD *, DWORD, DWORD *))
      GetProcAddress(hPsApi, "EnumProcesses");

   pEnumProcessModules = (BOOL (WINAPI *)(HANDLE, HMODULE *,
      DWORD, LPDWORD)) GetProcAddress(hPsApi, "EnumProcessModules");

   pGetModuleFileNameExW = (DWORD (WINAPI *)(HANDLE, HMODULE,
      LPWSTR, DWORD)) GetProcAddress(hPsApi, "GetModuleFileNameExW");

   if (pNtQuerySystemInformation   == NULL ||
      pNtQueryObject            == NULL ||
      pEnumProcesses            == NULL ||
      pEnumProcessModules         == NULL ||
      pGetModuleFileNameExW      == NULL)
      return FALSE;

   // winlogon position

   GetSystemDirectoryW(WinLogon, MAX_PATH * sizeof (WCHAR));
   wcscat(WinLogon, L"\\winlogon.exe");

   // set privileges

   if (SetPrivileges() == FALSE)
      return FALSE;

   // search winlogon

   dwSize2 = 256 * sizeof(DWORD);

   do
   {
      if (lpdwPIDs)
      {
         HeapFree(GetProcessHeap(), 0, lpdwPIDs);
         dwSize2 *= 2;
      }
      
      lpdwPIDs = (LPDWORD) HeapAlloc(GetProcessHeap(), 0, dwSize2);

        if (lpdwPIDs == NULL)
         return FALSE;
           
      if (!pEnumProcesses(lpdwPIDs, dwSize2, &dwSize))
         return FALSE;

   } while (dwSize == dwSize2);

   dwSize /= sizeof(DWORD);

   for (dwIndex = 0; dwIndex < dwSize; dwIndex++)
   {
      Buffer[0] = 0;
      
      hProcess = OpenProcess(PROCESS_QUERY_INFORMATION |
         PROCESS_VM_READ, FALSE, lpdwPIDs[dwIndex]);

      if (hProcess != NULL)
      {
         if (pEnumProcessModules(hProcess, &hMod,
            sizeof(hMod), &dwSize2))
         {
            if (!pGetModuleFileNameExW(hProcess, hMod,
               Buffer, sizeof(Buffer)))
            {
               CloseHandle(hProcess);
               continue;
            }
            }
         else
         {
            CloseHandle(hProcess);
            continue;
         }

         if (Buffer[0] != 0)
         {
            GetFileName(Buffer);
            
            if (CompareStringW(0, NORM_IGNORECASE,
               Buffer, -1, WinLogon, -1) == CSTR_EQUAL)
            {
               // winlogon process found
               WinLogonId = lpdwPIDs[dwIndex];
               CloseHandle(hProcess);
               break;
            }

            dwLIndex++;
         }

         CloseHandle(hProcess);
      }   

   }

   if (lpdwPIDs)
      HeapFree(GetProcessHeap(), 0, lpdwPIDs);

   hWinLogon = OpenProcess(PROCESS_DUP_HANDLE, 0, WinLogonId);

   if (hWinLogon == NULL)
   {
      return FALSE;
   }
   
   nt = pNtQuerySystemInformation(SystemHandleInformation, NULL, 0, &uSize);

   while (nt == STATUS_INFO_LENGTH_MISMATCH)
   {  
      uSize += 0x1000;
      
      if (pSystemHandleInfo)
         VirtualFree(pSystemHandleInfo, 0, MEM_RELEASE);
      
      pSystemHandleInfo = (PSYSTEMHANDLEINFO) VirtualAlloc(NULL, uSize,
         MEM_COMMIT, PAGE_READWRITE);
      
      if (pSystemHandleInfo == NULL)
      {
         CloseHandle(hWinLogon);
         return FALSE;
      }
         
      nt = pNtQuerySystemInformation(SystemHandleInformation,
         pSystemHandleInfo, uSize, &uBuff);
   }
   
   if (nt != STATUS_SUCCESS)
   {
      VirtualFree(pSystemHandleInfo, 0, MEM_RELEASE);
      CloseHandle(hWinLogon);
      return FALSE;
   }
   
   int handlesClosed = 0;

   for (i = 0; i < pSystemHandleInfo->nHandleEntries; i++)
   {
      if (pSystemHandleInfo->HandleInfo[i].Pid == WinLogonId)
      {
         if (DuplicateHandle(hWinLogon,
            (HANDLE) pSystemHandleInfo->HandleInfo[i].HandleValue,
            GetCurrentProcess(), &hCopy, 0, FALSE, DUPLICATE_SAME_ACCESS))
         {
            nt = pNtQueryObject(hCopy, ObjectNameInformation,
               &ObjName, sizeof (ObjName),NULL);
            
            if (nt == STATUS_SUCCESS)
            {
               wcsupr(ObjName.Buffer);

               if (CompareStringBackwards(ObjName.Buffer, L"WINDOWS\\SYSTEM32") ||
                  CompareStringBackwards(ObjName.Buffer, L"WINNT\\SYSTEM32"))
               {
                  // disable wfp on the fly
                  
                  CloseHandle(hCopy);
                     
                  DuplicateHandle (hWinLogon,
                     (HANDLE) pSystemHandleInfo->HandleInfo[i].HandleValue,
                     GetCurrentProcess(), &hCopy, 0, FALSE,
                     DUPLICATE_CLOSE_SOURCE | DUPLICATE_SAME_ACCESS);

                  CloseHandle(hCopy);
				  handlesClosed++;
               }
            }
            else
            {
               CloseHandle(hCopy);
            }

         }
      }
   }

   VirtualFree(pSystemHandleInfo, 0, MEM_RELEASE);
   CloseHandle(hWinLogon);

   printf("%d open SFC file handles closed\n", handlesClosed);

   if(!patchDllToo) return TRUE;

   // patch wfp smartly

   GetSystemDirectoryW(Buffer, sizeof (WCHAR) * MAX_PATH);
   GetSystemDirectoryW(Buffer2, sizeof (WCHAR) * MAX_PATH);

   wsprintfW(Buffer2, L"%s\\trash%X", Buffer2, GetTickCount());

   if (osvi.dwMajorVersion == 5 && osvi.dwMinorVersion == 0) // win2k
   {
      wcscat(Buffer, L"\\sfc.dll");
   }
   else // winxp, win2k3
   {
      wcscat(Buffer, L"\\sfc_os.dll");
   }

   hFile = CreateFileW(Buffer, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE,
      NULL, OPEN_EXISTING, 0, NULL);

   if (hFile == INVALID_HANDLE_VALUE)
   {
      return FALSE;
   }

   dwFileSize = GetFileSize(hFile, NULL);

   pSfc = (BYTE *) VirtualAlloc(NULL, dwFileSize, MEM_COMMIT, PAGE_READWRITE);

   if (!pSfc)
   {
      CloseHandle(hFile);
      return FALSE;
   }

   if (!ReadFile(hFile, pSfc, dwFileSize, &BRW, NULL))
   {
      CloseHandle(hFile);
      VirtualFree(pSfc, 0, MEM_RELEASE);
      return FALSE;
   }

   CloseHandle(hFile);

   ImgDosHeader = (PIMAGE_DOS_HEADER) pSfc;
   ImgNtHeaders = (PIMAGE_NT_HEADERS)
      (ImgDosHeader->e_lfanew + (ULONG_PTR) pSfc);
   ImgSectionHeader = IMAGE_FIRST_SECTION(ImgNtHeaders);

   // code section

   pCode = (BYTE *) (ImgSectionHeader->PointerToRawData + (ULONG_PTR) pSfc);

   // i gotta find the bytes to patch

   for (dwCount = 0; dwCount < (ImgSectionHeader->SizeOfRawData - 10); dwCount++)
   {
      if (pCode[dwCount] == 0x8B && pCode[dwCount + 1] == 0xC6 &&
         pCode[dwCount + 2] == 0xA3 && pCode[dwCount + 7] == 0x3B &&
         pCode[dwCount + 9] == 0x74 && pCode[dwCount + 11] == 0x3B)
      {
         bFound = TRUE;
         break;
      }
   }

   if (bFound == FALSE)
   {
      // cannot patch
      // maybe w2k without sp1
	  printf("Could not find byte sequence to patch!\n");
      goto no_need_to_patch;
   }

   // patch

   pCode[dwCount] = pCode[dwCount + 1] = 0x90;

   // move dll to another place

   MoveFileW(Buffer, Buffer2);

   // create new dll

   hFile = CreateFileW(Buffer, GENERIC_WRITE, FILE_SHARE_READ,
      NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

   if (hFile == INVALID_HANDLE_VALUE)
   {
      // cannot patch

      VirtualFree(pSfc, 0, MEM_RELEASE);
      return FALSE;
   }

   WriteFile(hFile, pSfc, dwFileSize, &BRW, NULL);

   CloseHandle(hFile);
	
   printf("SFC Dll Permanently Patched!\n");

no_need_to_patch:

   VirtualFree(pSfc, 0, MEM_RELEASE);

   // modify the registry

   Ret = RegOpenKeyExW(HKEY_LOCAL_MACHINE,
      L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon",
      0, KEY_SET_VALUE, &Key);

   if (Ret != ERROR_SUCCESS)
   {
      return FALSE;
   }

   BRW = 0xFFFFFF9D;

   Ret = RegSetValueExW(Key, L"SFCDisable", 0, REG_DWORD, (PBYTE) &BRW, sizeof (BRW));

   if (Ret != ERROR_SUCCESS)
   {
      return FALSE;
   }

   BRW = 0;

   Ret = RegSetValueExW(Key, L"SFCScan", 0, REG_DWORD, (PBYTE) &BRW, sizeof (BRW));

   if (Ret != ERROR_SUCCESS)
   {
      return FALSE;
   }
   
   RegCloseKey(Key);
   printf("SFC disabled in the registry!\n");

   return TRUE;
}

void main()
{

   bool doPatch = false;

   printf("This will disable windows file protection.\n");
   printf("This is dangerous for system health.\n");
   printf("Author: Daniel Pistelli ntcore.com - 2004\n");
   printf("-- EXPERT USE ONLY, USE AT YOUR OWN RISK!!--\n\n");
   printf("Choose an option:\n");
   printf(" 1) temporary until next reboot\n");
   printf(" 2) permanently patch the dll\n");
   printf(" 3) exit now\n");
   printf("Enter choice (1-3):");

   char opt = getch();
	
   if(opt=='2'){
	   printf("\nConfirm permanent disable of WFP: (Y/N)");
	   char y = tolower(getch());
	   if(y=='n') return;
	   doPatch = true;
   }
   else if(opt!='1'){
	   return;
   }

   printf("\n\nDisabling SFC..\n");

   if (TrickWFP(doPatch) == TRUE)
   {
      printf("Complete!\n");  
   }
   else
   {
	   printf("Failed :(\n");
   }

   printf("press any key to continue..");
   getch();


}

