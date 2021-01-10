PROCEDURE Lpu2Postgre
 IF MESSAGEBOX('унрхре хлонпрхпнбюрэ дюммше б POSTGRESQL?',4+32,'')=7
  RETURN 
 ENDIF 

 nHandl = SQLCONNECT("postgresql")
 IF nHandl <= 0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(3)), 16, 'Cannot make connection')
  RETURN 
 ENDIF
 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\people.dbf')
  MESSAGEBOX('нрясрябсер тюик '+UPPER(pBase+'\'+m.gcPeriod+'\people.dbf'),0+64,'')
  =SQLDISCONNECT(nHandl)
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\people', 'people', 'shar')>0
  IF USED('people')
   USE IN people
  ENDIF 
  =SQLDISCONNECT(nHandl)
  RETURN 
 ENDIF 
 
 WAIT "хлонпр..." WINDOW NOWAIT 
 
 SELECT people
 m.nrecs = 0 
 m.dbeg = SECONDS()
 SCAN 
  SCATTER MEMVAR 
  m.tip_p = tipp
  m.dr = STR(YEAR(m.dr),4)+'-'+PADL(MONTH(m.dr),2,'0')+'-'+PADL(DAY(m.dr),2,'0')

*  cmd01 = 'insert into data.people (period,tip_p,sn_pol,fam,im,ot,w,dr) '
*  cmd02 = 'values (?m.gcperiod,?m.tip_p,?m.sn_pol,?m.fam,?m.im,?m.ot,?m.w,?m.dr)'
*  cmdAll = cmd01+cmd02

  cmdAll = 'select data.addperson(?m.gcperiod,?m.tip_p,?m.sn_pol,?m.fam,?m.im,?m.ot,?m.w,?m.dr)'

  IF SQLEXEC(nHandl, cmdAll)!=-1
*   lnGoodRecs = lnGoodRecs + 1
  ENDIF 

  m.nrecs = m.nrecs + 1
  IF m.nrecs>999
*   EXIT 
  ENDIF 
  
 ENDSCAN 
 m.dend = SECONDS()
 USE IN people 
 WAIT CLEAR 
 
 =SQLDISCONNECT(nHandl)
 MESSAGEBOX("бпелъ напюанрйх: "+TRANSFORM(m.dend-m.dbeg,'999999999') +' ЯЕЙ.',0+64,'')

RETURN 