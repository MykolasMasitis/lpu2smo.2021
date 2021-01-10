PROCEDURE AllMpTyp
 IF MESSAGEBOX('ПЕРЕСЧИТАТЬ Mp, Typ и Vz?',4+32,'')=7
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
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  IF USED('tarifn')
   USE IN tarifn 
  ENDIF 
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\pilot', 'pilot', 'shar', 'lpu_id')>0
  IF USED('pilot')
   USE IN pilot 
  ENDIF 
  USE IN tarif 
  USE IN aisoms
  RETURN 
 ENDIF 
 
 SELECT aisoms 
 SCAN 
  m.mcod  = mcod 
  m.lpuid = lpuid
  m.IsPilot = IIF(SEEK(m.lpuid, 'pilot'), .T., .F.)
  m.IsStac  = IIF(VAL(SUBSTR(m.mcod,3,2))>40,.t.,.f.)

  IF !fso.FolderExists(pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   USE IN talon 
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  SELECT talon 
  SET RELATION TO sn_pol INTO people 
  SCAN 
  
  m.cod = cod
  m.ds  = ds
  m.otd = SUBSTR(otd,2,2)
  m.d_type = d_type
  m.ord = ord
  m.lpu_ord = lpu_ord

  m.IsDental = IsDental(m.cod, m.lpuid, m.mcod, m.ds)
   
  m.IsTpnR    = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r', .T., .F.)
  m.Is02      = IIF(SEEK(m.cod, 'tarif') and tarif.tpn='q', .t., .f.)
  m.prmcod    = people.prmcod
  m.prmcods   = people.prmcods
  
  m.profil = profil

  IF m.IsDental

  DO CASE 
   CASE EMPTY(m.prmcods) && неприкрепленные
    m.Typ = '0'
    m.Mp = 's'
   
   CASE m.mcod  = m.prmcods && свои пациенты
    m.Typ = '1'
    DO CASE 
     CASE m.IsTpnR = .T. OR INLIST(m.otd,'08') && tpn='r' - 3 услуги по июлю 2019, 08 - 4
      m.Mp = '8'
     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod) && 23 услуги
      m.Mp = '8'
     CASE m.otd='93' AND IsStac(m.mcod) && ни одной!
      m.Mp = '8'
     OTHERWISE 
       m.Mp = 's'
    ENDCASE 
    
   CASE m.mcod != m.prmcods && чужие пациенты
    m.Typ = '2'
    m.Mp = 's'

   OTHERWISE 

  ENDCASE 

  ELSE && IF IsDental(m.cod, m.lpuid, m.mcod, m.ds)

  DO CASE 
   CASE EMPTY(m.prmcod) && неприкрепленные
    m.Typ = '0'
    DO CASE 
     CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08')) && Добавление условия pilot ничего не меняет
      m.Mp = '4'
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      m.Mp = 'm'
     CASE INLIST(m.otd,'01') AND IsStac(m.mcod)
      m.Mp = '4'
     CASE INLIST(m.otd,'70','73','90','93') AND IsStac(m.mcod)
      m.Mp = '4'
     CASE m.ord=7 AND m.lpu_ord=7665
      m.Mp = '4'
     CASE INLIST(INT(m.cod/1000),49,149) AND m.mcod!=people.prmcod AND people.tip_p=3 
      m.Mp = '4'
     CASE INLIST(INT(m.cod/1000),29,129) AND m.mcod!=people.prmcod AND people.tip_p=3 
      m.Mp = '4'
     OTHERWISE 
       m.Mp = 'p'
    ENDCASE 
   
   CASE m.mcod  = m.prmcod && свои пациенты
    m.Typ = '1'
    DO CASE 
     CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08'))
      m.Mp = '4'
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      m.Mp = 'm'
     CASE INLIST(m.otd,'70','73','93') AND IsStac(m.mcod)
      m.Mp = '4'
     OTHERWISE 
       m.Mp = 'p'
    ENDCASE 
    
   CASE m.mcod != m.prmcod && чужие пациенты
    m.Typ = '2'
    DO CASE 
     CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08'))
      m.Mp = '4'
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      m.Mp = 'm'
     CASE INLIST(m.otd,'01') AND IsStac(m.mcod)
      m.Mp = '4'
     CASE INLIST(m.otd,'70','73','90','93') AND IsStac(m.mcod)
      m.Mp = '4'
     CASE m.ord=7 AND m.lpu_ord=7665
      m.Mp = '4'
     CASE INLIST(INT(m.cod/1000),49,149) AND m.mcod!=people.prmcod AND people.tip_p=3 
      m.Mp = '4'
     CASE INLIST(INT(m.cod/1000),29,129) AND m.mcod!=people.prmcod AND people.tip_p=3 
      m.Mp = '4'
     OTHERWISE 
      m.Mp = 'p'
      DO CASE 
       CASE m.lpu_ord>0 && vz=1, направление, в т.ч. договор с ДШО/ШО, договор на проведение вакцинопрофилактики и "актив" ССиНМП
        m.vz = 1
       CASE m.Is02 && vz=2, неотложная помощь (по реестру медицинских услуг)
        m.vz = 2
       CASE m.profil='100' AND INLIST(m.otd,'00','92') && vz=3, услуги, оказанные в травмапункте (в дополнение к  коду 2)
        m.vz = 3
       CASE m.otd='08' && vz=4, услуги ЖК
        m.vz = 4
       CASE m.otd='91' && vz=5, услуги ЦЗ
        m.vz = 5
       OTHERWISE 
        m.vz = 0 && то, что должно попасть в up-файл
      ENDCASE 
    ENDCASE 
   OTHERWISE 
  ENDCASE 

  ENDIF IF IsDental(m.cod, m.lpuid, m.mcod, m.ds)

  ENDSCAN 
  SET RELATION OFF INTO people
  USE 
  USE IN people
  SELECT aisoms
  
 ENDSCAN 
 USE
 
 USE IN tarif 
 USE IN pilot 
 
 MESSAGEBOX('OK!',0+64,'')

RETURN 