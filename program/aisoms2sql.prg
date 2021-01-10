PROCEDURE AisOms2sql

 IF MESSAGEBOX('хлонпрхпнбюрэ ад б SQL?',4+32,'')=7
  RETURN 
 ENDIF 

 IF !fso.FolderExists(pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\aisoms.dbf')
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
 
 DO OpenBases
 
 m.d_beg = SECONDS()
 
 DO aisoms

 m.d_end = SECONDS()
 WAIT CLEAR 
 
 IF USED('transid')
  USE IN transid
 ENDIF 
 
 MESSAGEBOX("бпелъ напюанрйх: "+TRANSFORM(m.d_end-m.d_beg,'999999999') +' ЯЕЙ.',0+64,'')

 =SQLDISCONNECT(nHandl)

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

PROCEDURE AisOms
  IF USED('aisoms')
   lnGoodRecs = 0
   SELECT aisoms 
   SCAN 
    SCATTER MEMVAR 

    cmd01 = 'INSERT INTO dbo.Aisoms '
    cmd02 = '(period, lpuid, mcod, pazval, finval, pazvals, finvals, paz, nsch, s_pred, s_lek, '
    cmd03 = 's_mek, s_mek2, s_532, st_flk, pf_flk, s_ppa, ls_flk, s_avans, s_pr_avans, s_avans2, s_pr_avans2, e_mee, e_ekmp, dolg_b'
    cmd04 = ''
    cmd05 = ''
    cmd06 = ')'
    cmd07 = 'VALUES '
    cmd08 = '(?m.gcperiod, ?m.lpuid, ?m.mcod, ?m.pazval, ?m.finval, ?m.pazvals, ?m.finvals, ?m.paz, ?m.nsch, ?m.s_pred, ?m.s_lek, '
    cmd09 = '?m.sum_flk, ?m.sum_flk2, ?m.s_532, ?m.st_flk, ?m.pf_flk, ?m.s_ppa, ?m.ls_flk, ?m.s_avans, ?m.s_pr_avans, ?m.s_avans2, ?m.pr_avans2, ?m.e_mee, ?m.e_ekmp, ?m.dolg_b'
    cmd10 = ''
    cmd11 = ''
    cmd12 = ')'
    cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10+cmd11+cmd12
   
    IF SQLEXEC(nHandl, cmdAll)!=-1
     lnGoodRecs = lnGoodRecs + 1
    ELSE 
     =AERROR(errarr)
     =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'Cannot load data')
     =MESSAGEBOX(ALLTRIM(errarr(3)), 16, 'Cannot load data')
     EXIT 
    ENDIF 

   ENDSCAN 
   USE 
 ENDIF 
RETURN 