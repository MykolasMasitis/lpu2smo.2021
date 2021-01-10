PROCEDURE SumAnswers
 IF MESSAGEBOX('СОБРАТЬ ANSWERS?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms 
  ENDIF 
  RETURN 
 ENDIF 
 
 IF fso.FileExists(pBase+'\'+gcPeriod+'\allans.dbf')
  =OpenFile(pBase+'\'+gcPeriod+'\allans', 'ans', 'shar')
  *=OpenFile(pbase+'\'+m.gcperiod+'\nsi\outs', 'enp', 'shar', 'enp')
  *=OpenFile(pbase+'\'+m.gcperiod+'\nsi\outs', 'kms', 'shar', 'kms', 'again')
  *=OpenFile(pbase+'\'+m.gcperiod+'\nsi\outs', 'vsn', 'shar', 'vsn', 'again')
  =OpenFile(pbase+'\'+'202009'+'\nsi\outs', 'enp', 'shar', 'enp')
  =OpenFile(pbase+'\'+'202009'+'\nsi\outs', 'kms', 'shar', 'kms', 'again')
  =OpenFile(pbase+'\'+'202009'+'\nsi\outs', 'vsn', 'shar', 'vsn', 'again')
  
  SELECT ans 
  SCAN 
   m.tip_d = tip_d
   *m.q08=""
   m.q09=""
   m.n_pol=n_pol

   DO CASE 
    CASE m.tip_d='В'
     m.polis = ALLTRIM(m.n_pol)
     IF LEN(m.polis)=9
      *m.q08 = IIF(SEEK(m.polis, 'vsn'), vsn.q, "")
      m.q09 = IIF(SEEK(m.polis, 'vsn'), vsn.q, "")
     ENDIF 

    CASE INLIST(m.tip_d,'П','Э','К')
     m.polis   = LEFT(m.n_pol,16)
     *m.q08     = IIF(SEEK(m.polis, 'enp'), enp.q, "")
     m.q09     = IIF(SEEK(m.polis, 'enp'), enp.q, "")

    CASE m.tip_d='С'
     m.polis = ALLTRIM(m.n_pol)
     *m.q08     = IIF(SEEK(m.polis, 'kms'), kms.q, "")
     m.q09     = IIF(SEEK(m.polis, 'kms'), kms.q, "")

    OTHERWISE 
     && оставляем так, как подало МО
   ENDCASE 
   
   *REPLACE q08 WITH m.q08
   REPLACE q09 WITH m.q09
  
  ENDSCAN 
  USE 

  USE IN enp
  USE IN kms
  USE IN vsn
  
  RETURN 
 ENDIF 
 
 CREATE CURSOR curss (recid c(6), s_pol c(6), n_pol c(16), d_u c(8), q c(2), fam c(25), im c(20), ot c(20), ;
   dr c(8), w n(1), ans_r c(3), tip_d c(1), lpu_id n(6), st_id n(6), pd_id n(6), d_rq d, d_end d)
 INDEX on n_pol TAG n_pol

 SELECT aisoms 
 SCAN 
  m.mcod = mcod 
  IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\soapans.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\soapans', 'answers', 'shar')>0
   IF USED('answers')
    USE IN answers 
   ENDIF 
   SELECT aisoms 
   LOOP 
  ENDIF 
  
  WAIT m.mcod+"..." WINDOW NOWAIT 

  SELECT answers
  IF RECCOUNT('answers')=0
   USE IN answers 
   SELECT aisoms 
   LOOP 
  ENDIF 
  
  SCAN 
   SCATTER MEMVAR 
   IF SEEK(m.n_pol, 'curss')
    LOOP 
   ENDIF 
   INSERT INTO curss FROM MEMVAR 
  ENDSCAN 
  USE 
  
  SELECT aisoms 
  
  WAIT CLEAR 
  
 ENDSCAN
 USE IN aisoms
 
 SELECT curss
 COPY TO &pBase\&gcPeriod\allans WITH cdx 
 
 MESSAGEBOX('OK!',0+64,'')

RETURN 