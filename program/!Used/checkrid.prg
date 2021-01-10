PROCEDURE CheckRID
 
 IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–Œ¬≈–»“‹ '+CHR(13)+CHR(10)+;
               '“»œ RECID?'+CHR(13)+CHR(10)+;
               '',4+48,'') != 6
  RETURN 
 ENDIF 

 IF MESSAGEBOX('¬€ ¿¡—ŒÀﬁ“ÕŒ ”¬≈–≈Õ€ ¬ —¬Œ»’ ƒ≈…—“¬»ﬂ’?',4+48,'') != 6
  RETURN 
 ENDIF 

 ppriod = STR(tYear,4)+PADL(tMonth,2,'0')
 spriod = PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)

 ppdir  = pbase+'\'+ppriod
 IF !fso.FolderExists(ppdir)
  MESSAGEBOX('Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ '+ppdir,0+16,'')  
  RETURN
 ENDIF 
 
 aisfile = ppdir+'\AisOms'
 IF !fso.FileExists(aisfile+'.dbf')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À '+aisfile,0+16,'')  
  RETURN
 ENDIF 
 
 IF OpenFile(aisfile, 'AisOms', 'shared', 'mcod')>0
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\UsrLpu', "UsrLpu", "shar", "mcod") > 0
  USE IN aisoms
  RETURN
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\tarifn', "tarif", "shar", "cod") > 0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  USE IN aisoms
  RETURN
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\profus', "profus", "shar", "cod") > 0
  IF USED('profus')
   USE IN profus
  ENDIF 
  USE IN tarif
  USE IN aisoms
  RETURN
 ENDIF 
 IF OpenFile(pcommon+'\dspcodes', "dspcodes", "shar", "cod") > 0
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  USE IN profus
  USE IN tarif
  USE IN aisoms
  RETURN
 ENDIF 
 
 IF fso.FileExists(pbase+'\'+gcperiod+'\dsp.dbf')
  IF OpenFile(pbase+'\'+gcperiod+'\dsp', "dsp", "excl") > 0
   IF USED('dsp')
    USE IN dsp
   ENDIF 
  ELSE 
   SELECT dsp
   IF FIELD('tip')!='TIP'
    ALTER TABLE dsp ADD COLUMN Tip n(1)
   ENDIF 
   SET RELATION TO cod INTO dspcodes
    REPLACE ALL tip WITH dspcodes.tip
   SET RELATION OFF INTO dspcodes
   USE IN dspcodes 
   INDEX on mcod+sn_pol+PADL(tip,1,"0") TAG NewExpTag
   USE IN dsp 
  ENDIF 
 ENDIF 

 SELECT AisOms
 SCAN
  m.mcod = mcod
  m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
  m.lpuid = STR(lpuid,4)
  m.nvfile = 'nv'+m.lpuid

  WAIT m.mcod WINDOW NOWAIT 

  IF !fso.FolderExists(ppdir+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(ppdir+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(ppdir+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(ppdir+'\'+m.mcod+'\people', 'people', 'excl')>0
   LOOP 
  ENDIF 
  IF OpenFile(ppdir+'\'+m.mcod+'\talon', 'talon', 'excl')>0
   USE IN People
   LOOP 
  ENDIF 
  IF OpenFile(ppdir+'\'+m.mcod+'\m'+m.mcod, 'merror', 'excl')>0
   USE IN People
   LOOP 
  ENDIF 

  IF fso.FileExists(ppdir+'\'+m.mcod+'\people.bak')
   fso.DeleteFile(ppdir+'\'+m.mcod+'\people.bak')
  ENDIF 
  IF fso.FileExists(ppdir+'\'+m.mcod+'\talon.bak')
   fso.DeleteFile(ppdir+'\'+m.mcod+'\talon.bak')
  ENDIF 
  IF fso.FileExists(ppdir+'\'+m.mcod+'\otdel.bak')
   fso.DeleteFile(ppdir+'\'+m.mcod+'\otdel.bak')
  ENDIF 
  IF fso.FileExists(ppdir+'\'+m.mcod+'\doctor.bak')
   fso.DeleteFile(ppdir+'\'+m.mcod+'\doctor.bak')
  ENDIF 

  IF fso.FileExists(ppdir+'\'+m.mcod+'\ho'+m.qcod+'.dbf')
   IF OpenFile(ppdir+'\'+m.mcod+'\ho'+m.qcod, 'ho', 'excl')>0
    IF USED('ho')
     USE IN ho
    ENDIF 
   ELSE 
    SELECT ho
    IF FSIZE('c_i')!=30
     ALTER table ho alter COLUMN c_i c(30)
    ENDIF 
    USE IN ho 
   ENDIF 
  ENDIF 
  
  SELECT merror
  IF FIELD('subet')!='SUBET'
   ALTER TABLE merror ADD COLUMN SubEt n(1)
  ENDIF 
  IF FIELD('reason')!='REASON'
   ALTER TABLE merror ADD COLUMN reason c(1)
  ENDIF 
  IF FIELD('n_akt')!='N_AKT'
   ALTER TABLE merror ADD COLUMN n_akt c(15)
  ELSE 
   IF VARTYPE(n_akt) != 'C'
    ALTER TABLE merror drop COLUMN n_akt
    ALTER TABLE merror ADD COLUMN n_akt c(15)
   ENDIF 
  ENDIF 
  IF FIELD('d_akt')!='D_AKT'
   ALTER TABLE merror ADD COLUMN d_akt d
  ENDIF 
  IF FIELD('t_akt')!='T_AKT'
   ALTER TABLE merror ADD COLUMN t_akt c(2)
  ENDIF 
  IF FIELD('d_edit')!='D_EDIT'
   ALTER TABLE merror ADD COLUMN d_edit d
  ENDIF 
  
  SELECT talon 

  IF FSIZE('c_i')!=30
   ALTER TABLE talon ALTER COLUMN c_i c(30)
   DELETE TAG sn_pol
   INDEX on sn_pol FOR IsTlnValid() TAG sn_pol
  ENDIF 

  IF FIELD('MM')!='MM'
   ALTER TABLE Talon ADD COLUMN mm c(1)
  ENDIF 

  IF FIELD('f_type')!='F_TYPE'
   ALTER TABLE Talon ADD COLUMN f_type c(2)
  ELSE 
   IF FSIZE('f_type')!=2
    ALTER TABLE talon ALTER COLUMN f_type c(2)
   ENDIF 
  ENDIF 

  IF FIELD('Vz')!='VZ'
   ALTER TABLE Talon ADD COLUMN Vz L
  ENDIF 
  IF FIELD('lpu_ord')!='LPU_ORD'
   ALTER TABLE talon ADD COLUMN lpu_ord n(6)
  ENDIF 
  IF FIELD('date_ord')!='DATE_ORD'
   ALTER TABLE talon ADD COLUMN date_ord d
  ENDIF 
  IF FIELD('n_kd')!='N_KD'
   WAIT "ƒŒ¡¿¬À≈Õ»≈ N_KD..." WINDOW NOWAIT 
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
  IF FIELD('mp')!='MP'
   WAIT "ƒŒ¡¿¬À≈Õ»≈ MP..." WINDOW NOWAIT 
   ALTER TABLE talon ADD COLUMN mp c(1)
   WAIT CLEAR 
  ENDIF 
  
  IF m.qcod='S2'
  WAIT "œ–Œ¬≈– ¿ œ–Œ‘»Àﬂ..." WINDOW NOWAIT 
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
  ENDIF 
  
  SELECT people

  DELETE TAG sn_pol
  INDEX on sn_pol FOR IsPplValid() TAG sn_pol

  IF FIELD('prmcods')!='PRMCODS'
   ALTER TABLE People ADD COLUMN prmcods c(7)
  ENDIF 
  IF FIELD('IsPr')!='ISPR'
   ALTER TABLE People ADD COLUMN IsPr L
  ENDIF 
  IF FIELD('s_all')!='S_ALL'
   ALTER TABLE People ADD COLUMN s_all n(11,2)
  ENDIF 
  IF FIELD('fil_id')!='FIL_ID'
   ALTER TABLE People ADD COLUMN fil_id n(6)
  ENDIF 
  IF FIELD('prmcod')!='PRMCOD'
   ALTER TABLE People ADD COLUMN prmcod c(7)
  ENDIF 
  IF FIELD('tipp')!='TIPP'
   ALTER TABLE People ADD COLUMN tipp c(1)
   SCAN 
    DO CASE 
     CASE IsEnp(sn_pol)
      REPLACE tipp WITH 'œ'
     CASE IsKms(sn_pol)
      REPLACE tipp WITH '—'
     CASE IsVs(sn_pol)
      REPLACE tipp WITH '—'
     OTHERWISE 
      REPLACE tipp WITH '—'
    ENDCASE 
   ENDSCAN 
  ENDIF 
*  USE 

  IF !fso.FileExists(ppdir+'\'+m.mcod+'\'+m.nvfile+'.dbf')
   IF fso.FileExists(ppdir+'\'+m.mcod+'\'+m.nvfile+'.'+spriod)
    fso.CopyFile(ppdir+'\'+m.mcod+'\'+m.nvfile+'.'+spriod, ppdir+'\'+m.mcod+'\'+m.nvfile+'.dbf')
    oSettings.CodePage(ppdir+'\'+m.mcod+'\'+m.nvfile+'.dbf', 866, .t.)
    IF OpenFile(ppdir+'\'+m.mcod+'\'+nvfile, 'nvfile', 'excl') == 0
     SELECT nvfile 
     INDEX ON pcod TAG pcod 
     USE 
    ENDIF 
   ENDIF 
  ELSE 
*   fso.DeleteFile(ppdir+'\'+m.mcod+'\'+m.nvfile+'.dbf')
*   fso.DeleteFile(ppdir+'\'+m.mcod+'\'+m.nvfile+'.cdx')
  ENDIF 

  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('talon')
   USE IN talon 
  ENDIF 
  IF USED('merror')
   USE IN merror
  ENDIF 

  SELECT aisoms

 ENDSCAN 

 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 IF USED('usrlpu')
  USE IN UsrLpu
 ENDIF 
 IF USED('tarif')
  USE IN tarif
 ENDIF 
 IF USED('profus')
  USE IN profus
 ENDIF 
 IF USED('dspcodes')
  USE IN dspcodes
 ENDIF 
 
 WAIT CLEAR 

 MESSAGEBOX('OK!', 0+64, '')

RETURN 