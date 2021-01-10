DEFINE CLASS PrlFox as Session OLEPUBLIC 
 PROCEDURE Init
  PUBLIC fso AS SCRIPTING.FileSystemObject
  fso = CREATEOBJECT('Scripting.FileSystemObject')
  SET LIBRARY TO vfpzip.fll
  
  

 ENDPROC 
 
 PROCEDURE Destroy 
  SET LIBRARY TO 
  RELEASE fso
 ENDPROC 
 
 PROCEDURE ZipDir(toCallback)
 
  tcDir = 'd:\lpu2smo\base\201904\0150712'
  IF !fso.FolderExists(tcDir)
   RETURN 
  ENDIF 


  ZipName = 'myArc.zip'

  IF fso.FileExists(tcDir+'\'+ZipName)
   fso.DeleteFile(tcDir+'\'+ZipName)
  ENDIF 


  ZipOpen(ZipName, tcDir+'\')
  ZipFolder(tcDir)
  ZipClose()
  
  *ttt = toCallback.cersion
  
  toCallback.docmd('messageb("Hi!",0+64,"")')
  
 ENDPROC 
ENDDEFINE 