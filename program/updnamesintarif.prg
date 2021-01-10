PROCEDURE UpdNamesInTarif
 IF MESSAGEBOX('ÎÁÍÎÂÈÒÜ ÍÀÈÌÅÍÎÂÀÍÈß ÓÑËÓÃ Â ÒÀÐÈÔÅ?',4+32,'')=7
  RETURN 
 ENDIF 

 pUpdDir = fso.GetParentFolderName(pbin)+'\UPDATE'
 IF !fso.FolderExists(pUpdDir)
  fso.CreateFolder(pUpdDir)
 ENDIF 

 SET DEFAULT TO (pUpdDir)
 csprfile = ''
 csprfile=GETFILE('dbf')
 IF EMPTY(csprfile)
  MESSAGEBOX(CHR(13)+CHR(10)+'ÂÛ ÍÈ×ÅÃÎ ÍÅ ÂÛÁÐÀËÈ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 ospr = fso.GetFile(csprfile)
 IF LOWER(LEFT(ospr.name,4)) != 'rees'
  MESSAGEBOX(CHR(13)+CHR(10)+'ÝÒÎ ÍÅ ÒÀÐÈÔÍÛÉ ÑÏÐÀÂÎ×ÍÈÊ!'+CHR(13)+CHR(10),0+16,'reesxxnn')
  RELEASE ospr 
  RETURN 
 ENDIF 
 
 oSettings.CodePage(csprfile, 866, .t.)
 
 IF OpenFile(csprfile, 'spr', 'excl')>0
  =Exit()
  RETURN 
 ENDIF 
 
 SELECT spr
 
 m.IsOK = .T.
 IF m.IsOK = .T.
  IF FIELD('cod')!='COD'
   m.IsOK = .F.
  ENDIF 
 ENDIF 
 IF m.IsOK = .T.
  IF FIELD('name')!='NAME' AND FIELD('namem')!='NAMEM'
   m.IsOK = .F.
  ENDIF 
 ENDIF 

 IF !m.IsOK
  MESSAGEBOX(CHR(13)+CHR(10)+'ÍÅÂÅÐÍÀß ÑÒÐÓÊÒÓÐÀ ÔÀÉËÀ '+UPPER(ospr.name)+'!',0+16,'')
  =Exit()
  RETURN 
 ENDIF 
 
 MESSAGEBOX(CHR(13)+CHR(10)+'ÔÀÉË ÏÐÀÂÈËÜÍÛÉ!'+CHR(13)+CHR(10),0+64,'')
 
 INDEX on cod tag cod 
 SET ORDER TO cod
 
 m.curmon = INT(VAL(SUBSTR(m.gcperiod,5,2)))
 FOR m.nmonth=1 TO m.curmon
  m.lcperiod = LEFT(m.gcperiod,4)+PADL(m.nmonth,2,'0')
  m.lcmonth  = PADL(m.nmonth,2,'0')

  IF !fso.FolderExists(pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FolderExists(pbase+'\'+m.lcperiod+'\nsi')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\nsi\tarifn.dbf')
   LOOP 
  ENDIF 

  fso.CopyFile(pbase+'\'+m.lcperiod+'\nsi\tarifn.dbf', pbase+'\'+m.lcperiod+'\nsi\tarifno.dbf')
  fso.CopyFile(pbase+'\'+m.lcperiod+'\nsi\tarifn.cdx', pbase+'\'+m.lcperiod+'\nsi\tarifno.cdx')
  
  IF OpenFile(pbase+'\'+m.lcperiod+'\nsi\tarifn', 'tarifn', 'shar')>0
   IF USED('tarifn')
    USE IN tarifn
   ENDIF 
   LOOP 
  ENDIF 

  MESSAGEBOX(m.lcperiod,0+64,'')
  
  SELECT tarifn
  SET RELATION TO cod INTO spr
  SCAN 
   IF EMPTY(spr.cod)
    LOOP 
   ENDIF 
   IF name=spr.name
    LOOP 
   ENDIF 
   
   REPLACE name WITH spr.name
   
  ENDSCAN 
  SET RELATION OFF INTO spr 
  USE IN tarifn 
  
 ENDFOR 
 
 SELECT spr 
 SET ORDER TO 
 DELETE TAG ALL  
 USE IN spr
 
 MESSAGEBOX('OK!',0+64,'')

RETURN 

FUNCTION exit 
 IF USED('spr')
  USE IN spr 
 ENDIF 
 RELEASE ospr  
RETURN 
