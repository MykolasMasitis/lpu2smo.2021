PROCEDURE MakeUPFilesN
 IF MESSAGEBOX(CHR(13)+	CHR(10)+'бш унрхре ятнплхпнбюрэ UP-тюикш?'+CHR(13)+CHR(10),4+32,'мнбюъ бепяхъ')=7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(ptempl+'\upqqxxxx.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрярсрярбсер ьюакнм UPQQXXXX.DBF!'+CHR(13)+CHR(10),0+16,'Templates')
  RETURN 
 ENDIF 

 IF !fso.FolderExists(pbase+'\'+m.gcperiod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  USE IN tarif
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar', 'mcod')>0
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  USE IN sprlpu
  USE IN tarif
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilots', 'pilots', 'shar', 'mcod')>0
  IF USED('pilots')
   USE IN pilots
  ENDIF 
  USE IN pilot
  USE IN sprlpu
  USE IN tarif
  USE IN aisoms
  RETURN 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\pr4.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\pr4', 'pr4', 'shar', 'mcod')>0
   IF USED('pr4')
    USE IN pr4
   ENDIF 
  ENDIF 
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\pr4st.dbf') 
  IF OpenFile(pbase+'\'+m.gcperiod+'\pr4st', 'pr4st', 'shar', 'mcod')>0
   IF USED('pr4st')
    USE IN pr4st
   ENDIF 
  ENDIF 
 ENDIF 
 
 SELECT aisoms
 SCAN FOR !DELETED()

  m.lpuid = lpuid
  m.mcod  = mcod 

  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 

  m.IsPilot  = SEEK(m.mcod, 'pilot')
  m.IsPilotS = SEEK(m.mcod, 'pilots')

  *IF !m.IsPilot AND !m.IsPilotS
  IF !m.IsPilot
   LOOP 
  ENDIF 

  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('people')
    USE IN people 
   ENDIF 
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'error', 'shar', 'rid')>0
   IF USED('error')
    USE IN error 
   ENDIF 
   IF USED('people')
    USE IN people 
   ENDIF 
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 

  m.upfile = 'UP'+UPPER(m.qcod)+PADL(m.lpuid,4,'0')
  IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.upfile+'.dbf')
   fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.upfile+'.dbf')
  ENDIF 
  fso.CopyFile(ptempl+'\upqqxxxx.dbf', pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.upfile+'.dbf')
  =OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.upfile, 'upfile', 'shar')
  
  SELECT talon 
  SET RELATION TO recid INTO error
  SET RELATION TO sn_pol INTO people ADDITIVE 
  
  WAIT m.mcod WINDOW NOWAIT 
  
  m.t_sum = 0
  SCAN 
   IF !EMPTY(error.rid)
    LOOP 
   ENDIF 
   m.Mp  = Mp
   m.Typ = Typ
   IF !(m.Typ='2' AND INLIST(m.Mp,'p','s'))
    LOOP 
   ENDIF 
   m.vz = vz
   IF m.vz>0
    LOOP 
   ENDIF 
   
   m.prmcod  = people.prmcod
   m.prlpuid = IIF(SEEK(m.prmcod, 'sprlpu', 'mcod'), sprlpu.lpu_id, 0)
   m.cod     = cod
   m.k_u     = k_u
   m.s_all   = s_all

   m.t_sum = m.t_sum + m.s_all
   INSERT INTO upfile (period, lpu_id, cod, k_u, pr_all) VALUES ;
   	(m.gcperiod, m.prlpuid, m.cod, m.k_u, m.s_all)

  ENDSCAN 

  SELECT upfile
  REPLACE ALL recid WITH PADL(RECNO(),7,'0')
  USE IN upfile

  SELECT talon 
  SET RELATION OFF INTO people
  SET RELATION OFF INTO error
  USE IN people
  USE IN talon 
  USE IN error
  
  IF USED('pr4') AND USED('pr4st')
   m.pr4sum = 0
   m.pr4sum = m.pr4sum + IIF(SEEK(m.mcod, 'pr4'), pr4.s_bad, 0)
   m.pr4sum = m.pr4sum + IIF(SEEK(m.mcod, 'pr4st'), pr4st.s_bad, 0)
   IF m.t_sum != m.pr4sum
     MESSAGEBOX(CHR(13)+CHR(10)+'намюпсфемн пюяунфдемхе лефдс тюикнл PR4 х'+CHR(13)+CHR(10)+;
      'х ятнплхпнбюммшлх UP-тюикюлх!'+CHR(13)+CHR(10)+;
      'он дюммшл PR4(ST) ясллю S_BAD='+TRANSFORM(m.pr4sum,'99999999.99')+CHR(13)+CHR(10)+;
      'он дюммшл UP-тюикнб ясллю='+TRANSFORM(m.t_sum,'99999999.99')+CHR(13)+CHR(10), 0+48, m.mcod)
   ENDIF 
  ENDIF 
  
  WAIT CLEAR 
  
  SELECT aisoms
  
 ENDSCAN 
 USE IN aisoms
 USE IN sprlpu
 USE IN tarif
 USE IN pilot
 USE IN pilots
 IF USED('pr4')
  USE IN pr4
 ENDIF 
 IF USED('pr4st')
  USE IN pr4st
 ENDIF 
 
 MESSAGEBOX(CHR(13)+CHR(10)+'напюанрйю гюйнмвемю!'+CHR(13)+CHR(10),0+64,'')
 
RETURN 