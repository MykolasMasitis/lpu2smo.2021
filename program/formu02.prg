PROCEDURE FormU02
 IF MESSAGEBOX('—‘Œ–Ã»–Œ¬¿“‹ ‘Œ–Ã” ﬁ-2',4+32,'')=7
  RETURN 
 ENDIF 
* IF !fso.FileExists(pTempl+'\FormU01.xls')
*  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À ÿ¿¡ÀŒÕ¿ FormU02.xls',0+64,'')
*  RETURN 
* ENDIF 
 
 IF !fso.FolderExists(pbase+'\'+m.gcperiod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR curdata (nrec n(7), lpuid n(4), mcod c(7), ;
   prmcod c(7), prmcods c(7), period c(6), d_beg d, d_end d, s_all n(11,2), ;
   tip_p n(1), sn_pol c(25), tipp c(1), enp c(16), qq c(2), ;
   fam c(25), im c(20), ot c(20), w n(1), dr d, d_type c(1), ;
   sv c(3), recid_lpu c(7), c_err c(3))
 
 m.nrec = 0 

 SELECT aisoms
 SCAN 
  m.lpuid   = lpuid
  m.mcod    = mcod
*  m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')

  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'rerror', 'shar', 'rrid')>0
   USE IN people
   IF USED('rerror')
    USE IN rerror
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  m.nrec = m.nrec + 1

  SELECT people
  SET RELATION TO recid INTO rerror
  SCAN 
   IF EMPTY(rerror.c_err)
    LOOP 
   ENDIF 
   IF rerror.c_err='PNA'
    LOOP 
   ENDIF 

   SCATTER MEMVAR 
   m.c_err = rerror.c_err
   INSERT INTO curdata FROM MEMVAR 

  ENDSCAN 
  SET RELATION OFF INTO rerror
  USE 
  USE IN rerror

  SELECT aisoms

 ENDSCAN 
 USE IN aisoms 
 

 SELECT curdata 
 COPY TO &pbase\&gcperiod\FormU02

* m.llResult = X_Report(pTempl+'\FormU01.xls', pBase+'\'+m.gcperiod+'\FormU01.xls', .T.)

 USE 
 MESSAGEBOX('Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!',0+64,'')
 
RETURN 