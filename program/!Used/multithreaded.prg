FUNCTION MultiThreaded As VOID
	Declare Long CreateThreadWithObject in DMult.DLL ;
		String lpszClass, ;
		String lpszMethod, ;
		Object oRef,  ;
		Long @lpdwThreadId 
						
	Declare CloseHandle in Win32API LONG	

	LOCAL lnHandle, lnThreadID
	lnThreadID = 0
	
	lnHandle = CreateThreadWithObject( ;
					Strconv("mtdll.EasyMTServer"+Chr(0),5), ;
					Strconv("SomeLengthyProcess"+Chr(0),5), ;
					_VFP, ;
					@lnThreadID ;
				)
						
	=CloseHandle(m.lnHandle)		
ENDFUNC
