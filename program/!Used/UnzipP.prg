FUNCTION UnzipP(para1)
 PRIVATE m.fname
 m.fname = ALLTRIM(para1)
 
 fso = CREATEOBJECT('Scripting.FileSystemObject')
 IF !fso.FileExists(m.fname)
  RELEASE fso
  RETURN .F.
 ENDIF 
 
 SET LIBRARY TO d:\lpu2smo\bin\vfpzip.fll
 
 IF !UnZipOpen(m.fname)

  

RETURN 