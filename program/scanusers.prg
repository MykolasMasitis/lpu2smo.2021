PROCEDURE ScanUsers
 FOR nUser=1 TO 10
  lcUser = 'USR'+PADL(nUser,3,'0')
  IF fso.FolderExists(pAisOms+'\'+lcUser)
   WAIT '������������ ���������� ' + lcUser WINDOW NOWAIT 
   =CheckMail(lcUser)
   WAIT CLEAR 
  ENDIF 
 ENDFOR  
RETURN 