PROCEDURE CorSvBases
 IF MESSAGEBOX('��������� ������� ����?',4+32,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\talon.dbf')
  MESSAGEBOX('������� ���� �� �������!',0+16,'talon.dbf')
  RETURN 
 ENDIF 

 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\tarifn.dbf')
  MESSAGEBOX('����������� ���� TARIFN.DBF!',0+16,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pbase+'\'+m.gcperiod+'\talon', 'talon', 'excl')>0
  IF USED('talon')
   USE IN talon 
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\profus', 'profus', 'shar', 'cod')>0
  IF USED('profus')
   USE IN profus
  ENDIF 
  RETURN 
  USE IN talon 
  RETURN 
 ENDIF 
 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  USE IN profus
  RETURN 
  USE IN talon 
  RETURN 
 ENDIF 

 SELECT talon 
 WAIT "�������� �������..." WINDOW NOWAIT 
 SCAN 
  m.cod = cod 
  m.profil = profil
  IF EMPTY(m.profil)
   m.profil = IIF(SEEK(m.cod, 'profus'), ALLTRIM(profus.profil), '')
   REPLACE profil WITH m.profil
  ELSE 
   EXIT 
  ENDIF 
 ENDSCAN 
 WAIT CLEAR 

 IF FIELD('s_0')='S_0' AND FIELD('ds_0')!='DS_0'
  WAIT "�������� DS_0..." WINDOW NOWAIT 
  ALTER TABLE talon ADD COLUMN ds_0 c(6)
  ALTER TABLE talon DROP COLUMN s_0
  WAIT CLEAR 
 ENDIF 

 IF FIELD('fil_d')='FIL_D' AND FIELD('fil_id')!='FIL_ID'
  m.lIsFilId = .T.
  WAIT "�������� FIL_ID..." WINDOW NOWAIT 
  ALTER TABLE talon ADD COLUMN fil_id n(5)
  ALTER TABLE talon DROP COLUMN fil_d
  WAIT CLEAR 
  m.omcod = 'xxxxxxx'
  SCAN 
   m.mcod = mcod
   IF m.mcod!=m.omcod
    m.omcod = m.mcod
    IF USED('loctal')
     USE IN loctal
    ENDIF 
    IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'loctal', 'shar', 'recid')>0
     IF USED('loctal')
      USE IN loctal
     ENDIF 
     m.lIsFilId = .F.
     EXIT 
    ENDIF  
    WAIT m.mcod+'...' WINDOW NOWAIT 
   ENDIF 
   
   m.brid = brid 
   m.fil_id = IIF(SEEK(m.brid, 'loctal'), loctal.fil_id, 0)
   REPLACE fil_id WITH m.fil_id
   
  ENDSCAN 
  WAIT CLEAR 
  IF m.lIsFilId = .F.
   MESSAGEBOX('�� ������� �������� ���� FIL_ID!',0+48,'')
  ENDIF 
 ENDIF 

 IF FIELD('ds_2')='DS_2'
  WAIT "�������� DS_2..." WINDOW NOWAIT 
  ALTER TABLE talon DROP COLUMN ds_2
  WAIT CLEAR 
 ENDIF 
 IF FIELD('ds_3')='DS_3'
  WAIT "�������� DS_3..." WINDOW NOWAIT 
  ALTER TABLE talon DROP COLUMN ds_3
  WAIT CLEAR 
 ENDIF 
 IF FIELD('kur')='KUR'
  WAIT "�������� KUR..." WINDOW NOWAIT 
  ALTER TABLE talon DROP COLUMN kur
  WAIT CLEAR 
 ENDIF 
 IF FIELD('codnom')='CODNOM'
  WAIT "�������� CODNOM..." WINDOW NOWAIT 
  ALTER TABLE talon DROP COLUMN codnom
  WAIT CLEAR 
 ENDIF 
 IF FIELD('det')='DET'
  WAIT "�������� DET..." WINDOW NOWAIT 
  ALTER TABLE talon DROP COLUMN det
  WAIT CLEAR 
 ENDIF 
 IF FIELD('vnov_m')='VNOV_M'
  WAIT "�������� VNOV_M..." WINDOW NOWAIT 
  ALTER TABLE talon DROP COLUMN vnov_m
  WAIT CLEAR 
 ENDIF 
 IF FIELD('tipgr')='TIPGR'
  WAIT "�������� TIPGR..." WINDOW NOWAIT 
  ALTER TABLE talon DROP COLUMN tipgr
  WAIT CLEAR 
 ENDIF 
 IF FIELD('k2')='K2'
  WAIT "�������� K2..." WINDOW NOWAIT 
  ALTER TABLE talon DROP COLUMN k2
  WAIT CLEAR 
 ENDIF 
 IF FIELD('novor')='NOVOR'
  WAIT "�������� NOVOR..." WINDOW NOWAIT 
  ALTER TABLE talon DROP COLUMN novor
  WAIT CLEAR 
 ENDIF 
 IF FIELD('n_kd')!='N_KD'
  WAIT "���������� N_KD..." WINDOW NOWAIT 
  ALTER TABLE talon ADD COLUMN n_kd n(3)
  SCAN 
   m.tip = tip 
   IF EMPTY(m.tip)
    LOOP 
   ENDIF 
   m.cod = cod 
   IF !SEEK(m.cod, 'tarif')
    LOOP 
   ENDIF 
   m.n_kd = tarif.n_kd
   REPLACE n_kd WITH m.n_kd
  ENDSCAN 
  WAIT CLEAR 
 ENDIF 

 WAIT "��������������..." WINDOW NOWAIT 
 DELETE TAG ALL 
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON brid TAG brid
 INDEX ON c_i TAG c_i
 INDEX ON sn_pol TAG sn_pol
 INDEX ON otd TAG otd
 INDEX ON ds TAG ds
 INDEX ON d_u TAG d_u
 INDEX ON cod TAG cod
 INDEX ON profil TAG profil
 WAIT CLEAR 
 
 USE IN talon 
 USE IN profus
 USE IN tarif
 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\talon.bak')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\talon.bak')
 ENDIF 

 MESSAGEBOX('OK!',0+64,'')

RETURN 