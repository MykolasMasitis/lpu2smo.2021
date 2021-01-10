PROCEDURE Pr4st2Sql
 IF MESSAGEBOX('ÈÌÏÎÐÒÈÐÎÂÀÒÜ PR4ST Â SQL?',4+32,'')=7
  RETURN 
 ENDIF 

 IF !fso.FolderExists(pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\pr4st.dbf')
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\pr4st', 'pr4', 'shar')>0
  IF USED('pr4')
   USE IN pr4
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
 
 SELECT pr4
 SCAN 
  SCATTER MEMVAR 

  cmd01 = 'INSERT INTO dbo.Pr4st '
  cmd02 = '(period, lpuid, mcod, pazval, finval, p_all, s_all, p_empty, s_empty, p_own, s_own, '
  cmd03 = 'p_guests, s_guests, p_others, s_others, p_bad, s_bad'
  cmd04 = ''
  cmd05 = ''
  cmd06 = ')'
  cmd07 = 'VALUES '
  cmd08 = '(?m.gcperiod, ?m.lpuid, ?m.mcod, ?m.pazval, ?m.finval, ?m.paz_all, ?m.s_all, ?m.paz_empty, ?m.s_empty, ?m.paz_own, ?m.s_own, '
  cmd09 = '?m.paz_guests, ?m.s_guests, ?m.paz_others, ?m.s_others, ?m.paz_bad, ?m.s_bad'
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
 USE 

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
