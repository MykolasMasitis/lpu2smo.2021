PROCEDURE Mag2sql

 IF MESSAGEBOX('ÈÌÏÎÐÒÈÐÎÂÀÒÜ MAG02 Â SQL?',4+32,'')=7
  RETURN 
 ENDIF 

 IF !fso.FolderExists(pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\FormMag02.dbf')
  RETURN 
 ENDIF 
 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\FormMag02', 'mag02', 'shar')>0
  IF USED('mag02')
   USE IN mag02
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
 
 m.d_beg = SECONDS()
 
 DO mag02

 m.d_end = SECONDS()
 WAIT CLEAR 
 
 MESSAGEBOX("ÂÐÅÌß ÎÁÐÀÁÎÒÊÈ: "+TRANSFORM(m.d_end-m.d_beg,'999999999') +' ñåê.',0+64,'')

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

PROCEDURE mag02
  IF USED('mag02')
   lnGoodRecs = 0
   SELECT mag02
   SCAN 
    SCATTER MEMVAR 
    * col06 - âñåãî ïî ïðîòîîëó

    cmd01 = 'INSERT INTO dbo.mag02 '
    cmd02 = '(period, lpuid, mcod, s_all, s_stac, s_dental, s_dental_0, s_dop_all, s_dop_dent, '
    cmd03 = 'mek_all, mek_p, mek_r, mek_dent_p, mek_dent_r, mek_dent_r_0'
    cmd04 = ''
    cmd05 = ''
    cmd06 = ')'
    cmd07 = 'VALUES '
    cmd08 = '(?m.gcperiod, ?m.lpuid, ?m.mcod, ?m.col06, ?m.col08, ?m.col09, ?m.col13, ?m.col08, ?m.col18, '
    cmd09 = '?m.col10, ?m.col14, ?m.col17, ?m.col15, ?m.col19, ?m.col16'
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