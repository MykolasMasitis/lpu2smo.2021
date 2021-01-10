PROCEDURE MakeUDFilesN
 IF MESSAGEBOX(CHR(13)+	CHR(10)+'�� ������ ������������ UD-�����?'+CHR(13)+CHR(10),4+32,'����� �������')=7
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
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar', 'lpu_id')>0
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\horlpu', 'horlpu', 'shar', 'lpu_id')>0
  IF USED('horlpu')
   USE IN horlpu
  ENDIF 
  USE IN pilot
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  USE IN horlpu
  USE IN pilot
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  USE IN sprlpu
  USE IN horlpu
  USE IN pilot
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\stpilot', 'stpilot', 'shar', 'lpu_id')>0
  IF USED('stpilot')
   USE IN stpilot
  ENDIF 
  USE IN tarif
  USE IN sprlpu
  USE IN horlpu
  USE IN pilot
  USE IN aisoms
  RETURN 
 ENDIF 

 
 dufile = pbase+'\'+m.gcperiod+'\ud'+m.qcod+m.gcperiod
 IF fso.FileExists(dufile + '.dbf')
  fso.DeleteFile(dufile + '.dbf')
 ENDIF 
 
 CREATE TABLE &dufile ;
  (mcod c(7), lpu_id n(6), prmcod c(7), prlpu_id n(6), sn_pol c(25), cod n(6), k_u n(3), pr_all n(12,2), vz n(1))
 USE 
 =OpenFile(dufile, 'dufile', 'shar')
 
 SELECT aisoms
 SCAN FOR !DELETED()

  m.lpuid    = lpuid
  m.mcod     = mcod 
  m.IsStomat = IIF(SUBSTR(m.mcod,3,2)='07', .T., .F.)
  m.IsIskl   = IIF(INLIST(m.lpuid, 1912, 1940, 2049), .T., .F.)
  
  m.IsPilot  = SEEK(m.lpuid, 'pilot')
  m.IsHorLpu = SEEK(m.lpuid, 'horlpu')
  
  IF !m.IsPilot && ������� 26.11.2018!
   IF !SEEK(m.lpuid, 'horlpu')
    * LOOP  && ������� ��� �������������
   ENDIF 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
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

  SELECT talon 
  SET RELATION TO recid INTO error
  SET RELATION TO sn_pol INTO people ADDITIVE 
  
  WAIT m.mcod WINDOW NOWAIT 
  
  SCAN 
   IF !EMPTY(error.rid)
    LOOP 
   ENDIF 
   m.Mp  = Mp
   m.Typ = Typ
   IF !(m.Typ='2' AND m.Mp='p')
    LOOP 
   ENDIF 
   
   m.vz = vz
   IF m.vz<=0
    LOOP 
   ENDIF 
   IF m.vz=4
    *LOOP 
   ENDIF 
   *IF m.vz=6
   * m.ord     = ord
   * m.lpu_ord = lpu_ord
   * IF !(INLIST(m.ord,2,3) OR m.lpu_ord=1989)
   *  LOOP 
   * ENDIF 
   *ENDIF 
   
   m.prmcod  = people.prmcod
   m.prlpuid = IIF(SEEK(m.prmcod, 'pilot', 'mcod'), pilot.lpu_id, 0)
   m.cod     = cod
   m.k_u     = k_u
   m.s_all   = s_all
   m.sn_pol  = sn_pol

   INSERT INTO dufile (mcod, lpu_id, prmcod, prlpu_id, cod, k_u, pr_all, sn_pol, vz) VALUES ;
   	(m.mcod, m.lpuid, m.prmcod, m.prlpuid, m.cod, m.k_u, m.s_all, m.sn_pol, m.vz)

  ENDSCAN 

  SET RELATION OFF INTO people
  SET RELATION OFF INTO error
  USE IN people
  USE IN talon 
  USE IN error
  
  WAIT CLEAR 
  
  SELECT aisoms
  
 ENDSCAN 
 USE IN aisoms
 
 SELECT dufile
 SELECT prmcod DISTINCT FROM dufile INTO CURSOR svcur
 IF _tally<=0
  USE IN svcur
  USE IN dufile
  MESSAGEBOX(CHR(13)+CHR(10)+'��������� ���������!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(ptempl+'\udqqxxxx.dbf')
  USE IN svcur
  USE IN dufile
  MESSAGEBOX(CHR(13)+CHR(10)+'������������ ������ UDQQXXXX.DBF!'+CHR(13)+CHR(10),0+16,'Templates')
  RETURN 
 ENDIF 
 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\pr4.dbf') AND fso.FileExists(pbase+'\'+m.gcperiod+'\pr4st.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\pr4', 'pr4', 'shar', 'mcod')<=0 AND ;
  	OpenFile(pbase+'\'+m.gcperiod+'\pr4st', 'pr4st', 'shar', 'mcod')<=0
*   SELECT prmcod as mcod, SUM(pr_all) as s_all WHERE vz=1 FROM dufile GROUP BY prmcod INTO CURSOR curstat
   SELECT prmcod as mcod, SUM(pr_all) as s_all FROM dufile GROUP BY prmcod INTO CURSOR curstat
   SELECT curstat
   SET RELATION TO mcod INTO pr4
   SET RELATION TO mcod INTO pr4st ADDITIVE 
   m.sumstat = 0
   m.sumpr4  = 0
   SCAN 
    m.mcod = mcod
    *IF s_all != pr4.s_others + pr4st.s_others
    IF s_all != pr4.s_others
     *MESSAGEBOX(CHR(13)+CHR(10)+'���������� ����������� ����� ������ PR4 �'+CHR(13)+CHR(10)+;
      '� ��������������� UD-�������!'+CHR(13)+CHR(10)+;
      '�� ������ PR4 ����� S_OTHERS='+TRANSFORM(pr4.s_others + pr4st.s_others,'99999999.99')+CHR(13)+CHR(10)+;
      '�� ������ UD-������ �����='+TRANSFORM(s_all,'99999999.99')+CHR(13)+CHR(10),0+48, m.mcod)
     *IF m.mcod='0343003'
     * REPLACE pr4.s_others WITH s_all IN pr4 FOR mcod=m.mcod
     *ELSE 
      MESSAGEBOX(CHR(13)+CHR(10)+'���������� ����������� ����� ������ PR4 �'+CHR(13)+CHR(10)+;
      	'� ��������������� UD-�������!'+CHR(13)+CHR(10)+;
      	'�� ������ PR4 ����� S_OTHERS='+TRANSFORM(pr4.s_others,'9999999999.99')+CHR(13)+CHR(10)+;
      	'�� ������ UD-������ �����='+TRANSFORM(s_all,'9999999999.99')+CHR(13)+CHR(10),0+48, m.mcod)
     *ENDIF 
    ENDIF 
    m.sumstat = m.sumstat + s_all
    *m.sumpr4  = m.sumpr4 + pr4.s_others + pr4st.s_others
    m.sumpr4  = m.sumpr4 + pr4.s_others
   ENDSCAN 
   SET RELATION OFF INTO pr4
   SET RELATION OFF INTO pr4st
   USE IN pr4
   USE IN pr4st
   USE IN curstat
   IF m.sumstat!=m.sumpr4
    MESSAGEBOX(CHR(13)+CHR(10)+'���������� ����������� ����� ������ PR4 �'+CHR(13)+CHR(10)+;
     '� ��������������� UD-�������!'+CHR(13)+CHR(10)+;
     '�� ������ PR4 ����� S_OTHERS='+TRANSFORM(m.sumpr4,'9999999999.99')+CHR(13)+CHR(10)+;
     '�� ������ UD-������ �����='+TRANSFORM(m.sumstat,'9999999999.99')+CHR(13)+CHR(10),0+48,'')
   ELSE 
    MESSAGEBOX(CHR(13)+CHR(10)+'����� S_OTHERS ����� PR4'+CHR(13)+CHR(10)+;
     '������������� ����� UD-������'+CHR(13)+CHR(10)+;
     '� ���������� '+TRANSFORM(m.sumstat,'9999999999.99')+CHR(13)+CHR(10),0+64,'')
   ENDIF 
  ELSE 
   IF USED('pr4')
    USE 
   ENDIF 
   MESSAGEBOX(CHR(13)+CHR(10)+'�� ������� ������� ���� PR4,'+CHR(13)+CHR(10)+;
    '������ ������ �� ������������!',0+64,'')
  ENDIF 
 ELSE 
  MESSAGEBOX(CHR(13)+CHR(10)+'���� PR4 �� �����������,'+CHR(13)+CHR(10)+;
   '������ ������ �� ������������!',0+64,'')
 ENDIF 
 
 SELECT svcur
 SCAN 
  m.prmcod  = prmcod 
  m.lppid  = IIF(SEEK(m.prmcod, 'pilot', 'mcod'), pilot.lpu_id, 0)
  *m.lppid  = IIF(m.lppid>0, m.lppid, IIF(SEEK(m.prmcod, 'pilots', 'mcod'), pilots.lpu_id, 0))
  IF m.lppid<=0
   LOOP
  ENDIF 
  m.uddfile = 'UD'+UPPER(m.qcod)+PADL(m.lppid,4,'0')
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.prmcod)
   LOOP 
  ENDIF 
  IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.prmcod+'\'+m.uddfile+'.dbf')
   fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.prmcod+'\'+m.uddfile+'.dbf')
  ENDIF 

  fso.CopyFile(ptempl+'\udqqxxxx.dbf', pbase+'\'+m.gcperiod+'\'+m.prmcod+'\'+m.uddfile+'.dbf')
  
  =OpenFile(pbase+'\'+m.gcperiod+'\'+m.prmcod+'\'+m.uddfile, 'udfile', 'shar')
  
  SELECT * FROM dufile WHERE prmcod = m.prmcod INTO CURSOR curmmm
  SELECT curmmm
  m.mcod   = mcod
  m.lpu_id = lpu_id
  m.period = m.gcperiod
  SCAN 
   m.cod    = cod 
   m.k_u    = k_u
   m.pr_all = pr_all
   m.lpu_id = lpu_id
   m.vz     = vz
   *IF m.vz=1
    INSERT INTO udfile FROM MEMVAR 
   *ENDIF 
  ENDSCAN 
  USE IN curmmm
  SELECT udfile
  REPLACE ALL recid WITH PADL(RECNO(),7,'0')
  USE IN udfile
  
  SELECT svcur
  
 ENDSCAN 
 USE IN svcur
 USE IN horlpu 
 USE IN pilot 
 USE IN dufile
 USE IN sprlpu
 USE IN tarif
 IF USED('stpilot')
  USE IN stpilot
 ENDIF 

 MESSAGEBOX(CHR(13)+CHR(10)+'��������� ���������!'+CHR(13)+CHR(10),0+64,'')
 
RETURN 