PROCEDURE Otdel2Sql
 IF MESSAGEBOX('ÈÌÏÎÐÒÈÐÎÂÀÒÜ PEOPLE Â SQL?',4+32,'')=7
  RETURN 
 ENDIF 

 IF !fso.FolderExists(pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 

 nHandl = SQLCONNECT("local")
 IF nHandl <= 0
  nHandl = SQLCONNECT("lpu", "sa", "admin")
 ENDIF 
 
 IF nHandl <= 0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(3)), 16, 'Cannot make connection')
  RETURN 
 ENDIF

 =SetSession()
 
 IF SQLEXEC(nHandl, "USE lpu") = -1
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(3)), 16, 'Cannot use lpu')
  m.lResult = .F.
 ENDIF 
 
 SELECT aisoms
 SCAN 
  m.mcod  = mcod 

  IF !fso.FolderExists(pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\otdel.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\otdel', 'otdel', 'shar')>0
   IF USED('otdel')
    USE IN otdel
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'..' WINDOW NOWAIT 
  
  SELECT otdel

  SCAN 
   SCATTER MEMVAR 

   cmd01 = 'INSERT INTO dim.departments '
   cmd02 = '(mcod, otd, name, cnt_bed'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = ')'
   cmd07 = 'VALUES '
   cmd08 = '(?m.mcod, ?m.iotd, ?m.name, ?m.cnt_bed'
   cmd09 = ''
   cmd10 = ''
   cmd11 = ''
   cmd12 = ')'
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10+cmd11+cmd12
   
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ELSE 
    *=AERROR(errarr)
    *=MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'people')
    *=MESSAGEBOX(ALLTRIM(errarr(3)), 16, 'people')
    *EXIT 
    LOOP 
   ENDIF 

  ENDSCAN 
  USE 
  
  WAIT CLEAR 
 
  SELECT aisoms 
 ENDSCAN 
 USE IN aisoms

 =SQLDISCONNECT(nHandl)
 MESSAGEBOX('OK!',0+64,'')
RETURN 

FUNCTION SetSession()
 IF SQLEXEC(nHandl, "SET QUOTED_IDENTIFIER ON")<=0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'SET QUOTED_IDENTIFIER ON')
  RETURN 
 ENDIF 
 IF SQLEXEC(nHandl, "SET CONCAT_NULL_YIELDS_NULL ON")<=0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'SET CONCAT_NULL_YIELDS_NULL ON')
  RETURN 
 ENDIF 
 IF SQLEXEC(nHandl, "SET ANSI_NULLS ON")<=0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'SET ANSI_NULLS ON')
  RETURN 
 ENDIF 
 IF SQLEXEC(nHandl, "SET ANSI_PADDING ON")<=0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'SET ANSI_PADDING ON')
  RETURN 
 ENDIF 
 IF SQLEXEC(nHandl, "SET ANSI_WARNINGS ON")<=0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'SET ANSI_WARNINGS ON')
  RETURN 
 ENDIF 
 IF SQLEXEC(nHandl, "SET NUMERIC_ROUNDABORT OFF")<=0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'SET NUMERIC_ROUNDABORT OFF')
  RETURN 
 ENDIF 
RETURN 
