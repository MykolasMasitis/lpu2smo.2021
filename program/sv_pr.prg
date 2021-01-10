FUNCTION sv_pr(para1)
 m.mcod = para1
 
 IF !USED('enp')
  RETURN 
 ENDIF 
 IF !USED('kms')
  RETURN 
 ENDIF 
 IF !USED('vsn')
  RETURN 
 ENDIF 
 
 oal = ALIAS()

 m.lused = .T.
 IF !USED('people')
  m.lused = .F.
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people.dbf')
   IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people', 'people', 'shar')>0
    IF USED('people')
     USE IN people 
    ENDIF 
    SELECT (oal)
    RETURN 
   ENDIF 
  ELSE 
   RETURN 
  ENDIF 
 ENDIF 
 
 m.nrecs = 0
 SELECT people
 SCAN 
  m.tip   = tipp 
  m.d_type = d_type
  
  IF m.d_type='9'
   m.prmcod  = ''
   m.prmcods = ''
  ELSE 
  DO CASE 
   CASE m.tip='В'
    m.polis = ALLTRIM(sn_pol)
    IF LEN(m.polis)=9
     m.lpuid   = IIF(SEEK(m.polis, 'vsn'), vsn.lpu_tera, 0)
     m.prmcod  = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.mcod, '')
     m.lpuids  = IIF(SEEK(m.polis, 'vsn'), vsn.lpu_stom, 0)
     m.prmcods = IIF(SEEK(m.lpuids, 'sprlpu'), sprlpu.mcod, '')
    ENDIF 

   CASE INLIST(m.tip,'П','Э','К')
    *m.polis = enp
    m.polis = LEFT(sn_pol,16)
    m.lpuid   = IIF(SEEK(m.polis, 'enp'), enp.lpu_tera, 0)
    m.prmcod  = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.mcod, '')
    m.lpuids = IIF(SEEK(m.polis, 'enp'), enp.lpu_stom, 0)
    m.prmcods = IIF(SEEK(m.lpuids, 'sprlpu'), sprlpu.mcod, '')

   CASE m.tip='С'
    m.polis = ALLTRIM(sn_pol)
    m.lpuid   = IIF(SEEK(m.polis, 'kms'), kms.lpu_tera, 0)
    m.prmcod  = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.mcod, '')
    m.lpuids = IIF(SEEK(m.polis, 'kms'), kms.lpu_stom, 0)
    m.prmcods = IIF(SEEK(m.lpuids, 'sprlpu'), sprlpu.mcod, '')

   OTHERWISE 
    && ничего не меняем!
  ENDCASE 
  ENDIF 
  
  m.o_prmcod  = prmcod
  m.o_prmcods = prmcods
  
  IF USED('svoutsrslt')
   IF m.o_prmcod<>m.prmcod OR m.o_prmcods<>m.prmcods
    INSERT INTO svoutsrslt (mcod, sn_pol, prmcod, prmcods, pr_new, prs_new) VALUES ;
   	 (m.mcod, m.polis, m.o_prmcod, m.o_prmcods, m.prmcod, m.prmcods)
   ENDIF 
  ENDIF 
  
  REPLACE prmcod WITH m.prmcod, prmcods WITH m.prmcods
  
  m.nrecs = m.nrecs + 1
  IF m.nrecs/100 = INT(m.nrecs/100)
  	WAIT 'Обработано '+STR(m.nrecs,6) + ' записей...' WINDOW NOWAIT 
  ENDIF 
  
 ENDSCAN 
 
 IF !m.lused
  IF USED('people')
   USE IN people 
  ENDIF 
 ENDIF 
 
 SELECT &oal

RETURN 