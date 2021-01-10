DEFINE CLASS EasyMTServer as Session OLEPUBLIC 
 PROCEDURE SomeLengthyProcess(toCallBack)
 
 DECLARE Sleep IN WIN32API Long
 
 LOCAL lnCount as Integer
 
 FOR lnCount = 1 TO 20
 	toCallBack.DoCmd("? + ALLTRIM(SYS(2015))")
 	Sleep(1000)
 ENDFOR 
 
 toCallBack.DoCmd("? + ALLTRIM(SYS(2015))")
 
 CLEAR DLLS "Sleep"
 
 ENDPROC 
 
ENDDEFINE 