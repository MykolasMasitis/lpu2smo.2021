FUNCTION MakePrZip(lcPath, cmcod, clpuid)

 m.mmy = PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)
 m.ctrl = 'ctrl'+m.mmy+'.dbf'
 m.prsp = 'prsp'+m.qcod+m.mmy+'.pdf'

 IF !fso.FileExists(lcpath+'\'+m.ctrl)
  RETURN .F.
 ENDIF 
 IF !fso.FileExists(lcpath+'\'+m.prsp)
  RETURN .F.
 ENDIF 
 
* ZipFile = 'pm'+clpuid+m.qcod+'.zip'
 ZipFile = 'E'+m.qcod+clpuid+'.'+m.mmy
 IF fso.FileExists(lcPath+'\'+ZipFile)
  fso.DeleteFile(lcPath+'\'+ZipFile)
 ENDIF 
 
 ZipOpen(lcPath+'\'+ZipFile)
 ZipFile(lcPath+'\'+m.ctrl)
 ZipFile(lcPath+'\'+m.prsp)
 ZipClose()
 
 IF !fso.FileExists(lcPath+'\'+ZipFile)
  RETURN .F.
 ENDIF 

RETURN  .T.