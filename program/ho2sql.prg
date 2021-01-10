PROCEDURE HO2Sql
 IF MESSAGEBOX('ÈÌÏÎÐÒÈÐÎÂÀÒÜ HO Â SQL?',4+32,'')=7
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
 IF OpenFile(pCommon+'\periods', 'periods', 'shar', 'period')>0
  IF USED('periods')
   USE IN periods
  ENDIF 
  USE IN aisoms 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\pilot', 'pilot', 'shar', 'lpu_id')>0
  USE IN aisoms
  USE IN periods
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\pilots', 'pilots', 'shar', 'lpu_id')>0
  USE IN aisoms
  USE IN periods
  IF USED('pilots')
   USE IN pilots
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\horlpu', 'horlpu', 'shar', 'lpu_id')>0
  USE IN aisoms
  USE IN periods
  USE IN pilot
  USE IN pilots
  IF USED('horlpu')
   USE IN horlpu
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\horlpus', 'horlpus', 'shar', 'lpu_id')>0
  USE IN aisoms
  USE IN periods
  USE IN pilot
  USE IN pilots
  USE IN horlpu
  IF USED('horlpus')
   USE IN horlpus
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
  m.lpuid = lpuid
  m.mcod  = mcod 
  m.period_id = IIF(SEEK(m.gcperiod, 'periods'), periods.id, 0)

  m.IsPilot  = IIF(SEEK(m.lpuid, 'pilot'), .T., .F.)
  m.IsPilotS = IIF(SEEK(m.lpuid, 'pilots'), .T., .F.)
  m.IsHorLpu = IIF(SEEK(m.lpuid, 'horlpu'), .T., .F.)
  m.IsHorLpuS = IIF(SEEK(m.lpuid, 'horlpus'), .T., .F.)

  IF !fso.FolderExists(pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ho'+m.qcod+'.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ho'+m.qcod, 'ho', 'shar')>0
   IF USED('ho')
    USE IN ho 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   USE IN ho 
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   USE IN ho 
   USE IN talon 
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'err', 'shar', 'rid')>0
   USE IN ho 
   USE IN talon 
   USE IN people
   IF USED('err')
    USE IN err
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'..' WINDOW NOWAIT 
  
  SELECT * FROM talon INTO CURSOR ttl READWRITE 
  SELECT ttl 
  INDEX on c_i + PADL(cod,6,'0') TAG unik
  SET ORDER TO unik
  USE IN talon
  
  SELECT ttl 
  SET RELATION TO sn_pol INTO people
  SET RELATION TO recid INTO err ADDITIVE 
  SELECT ho 
  SET RELATION TO c_i + PADL(cod,6,'0') INTO ttl 
  

  SCAN 
   SCATTER MEMVAR 
   m.s_id    = ttl.recid_lpu
   m.fil_id  = ttl.fil_id
   m.d_u     = ttl.d_u
   m.ds      = ttl.ds
   m.prmcod  = people.prmcod
   m.prmcods = people.prmcods
   
   m.ismek = IIF(!EMPTY(err.c_err), .T., .F.)

   m.sex = people.w
   m.adj = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
   m.ages = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)

   cmd01 = 'INSERT INTO dbo.surgeries '
   cmd02 = '(s_id, period_id, period, lpuid, mcod, fil_id, '
   cmd03 = 'ages, sex, prmcod, sn_pol, c_i, cod, ho, k_ho, d_u, ismek, ds'
   cmd04 = ''
   cmd05 = ''
   cmd06 = ')'
   cmd07 = 'VALUES '
   cmd08 = '(?m.s_id, ?m.period_id, ?m.gcperiod, ?m.lpuid, ?m.mcod, ?m.fil_id, '
   cmd09 = '?m.ages, ?m.sex, ?m.prmcod, ?m.sn_pol, ?m.c_i, ?m.cod, ?m.codho, ?m.k_ho, ?m.d_u, ?m.ismek, ?m.ds'
   cmd10 = ''
   cmd11 = ''
   cmd12 = ')'
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10+cmd11+cmd12
   
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ELSE 
    =AERROR(errarr)
    =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'pr4st')
    =MESSAGEBOX(ALLTRIM(errarr(3)), 16, 'pr4st')
    EXIT 
   ENDIF 

  ENDSCAN 
 
  SET RELATION OFF INTO ttl 
  USE 
  SELECT ttl 
  SET RELATION OFF INTO people 
  USE
  USE IN people 
  USE IN err 
  
  WAIT CLEAR 
 
  SELECT aisoms 
 ENDSCAN 
 USE IN aisoms
 USE IN periods 
 USE IN pilot
 USE IN pilots
 USE IN horlpu
 USE IN horlpus

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
