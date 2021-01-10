FUNCTION CreateLicences
 IF MESSAGEBOX('ÑÎÇÄÀÒÜ ÔÀÉËÛ ËÈÖÅÍÇÈÉ?',4+32,'')=7
  RETURN 
 ENDIF 
 
 LicDir = fso.GetParentFolderName(pcommon) + 'LICENCES'
 IF !fso.FolderExists(LicDir)
  fso.CreateFolder(LicDir)
 ELSE 
  IF MESSAGEBOX('ÄÈÐÅÊÒÎÐÈß '+LicDir+' ÓÆÅ ÑÎÇÄÀÍÀ!'+CHR(13)+CHR(10)+'ÂÛ ÕÎÒÈÒÅ Å¨ ÏÅÐÅÑÎÇÄÀÒÜ?',4+15,'')=7
   RETURN 
  ENDIF 
 ENDIF   
 
 IF !fso.FileExists(pcommon+'\lpudogs.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË-ÑÏÐÀÂÎ×ÍÈÊ LPUDOGS.DBF!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 

 IF !fso.FileExists(pcommon+'\prv002xx.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË-ÑÏÐÀÂÎ×ÍÈÊ PRV002xx.DBF!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 

 IF OpenFile(pcommon+'\lpudogs', 'lpudogs', 'shar')>0
  IF USED('lpudogs')
   USE IN lpudogs
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pcommon+'\prv002xx', 'profs', 'shar')>0
  IF USED('lpudogs')
   USE IN lpudogs
  ENDIF 
  IF USED('profs')
   USE IN profs
  ENDIF 
  RETURN 
 ENDIF 
 
 IF OpenFile(pBase+'\'+gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod') > 0
  RETURN
 ENDIF 
 
 SELECT AisOms
 
 SCAN
  m.lpu_id = lpuid
  m.mcod   = mcod
  WAIT m.mcod WINDOW NOWAIT 
  IF !fso.FolderExists(pBase+'\'+gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar') > 0
   IF USED('talon')
    USE IN talon
   ENDIF 
   LOOP 
  ENDIF 
  
  licdbf = LicDir+'\l'+STR(m.lpu_id,4)
  CREATE CURSOR tmpr (vp n(1), prv n(3))
  INDEX on STR(vp,1)+STR(prv,3) TAG unik
  SET ORDER TO unik
  CREATE TABLE &licdbf (v l, tip n(1), cod n(3)) 
  licalias = ALIAS()
  SELECT &licalias
  INDEX on STR(tip,1)+STR(cod,3) TAG unik
  SET ORDER TO unik

  SELECT talon 
  SCAN 
   m.v   = .T.
   m.cod = INT(VAL(profil))
   m.usl = cod
   DO CASE 
    CASE IsGsp(m.usl)
     m.tip = 3
    CASE IsDst(m.usl)
     m.tip = 2
    CASE IsPlk(m.usl)
     m.tip = 1
    OTHERWISE 
     m.tip = 0
   ENDCASE 
   
   m.vp  = m.tip
   m.prv = m.cod
   
   m.vir = STR(m.tip,1)+STR(m.cod,3)
*   IF !SEEK(m.vir, '&licalias')
*    INSERT INTO &licalias FROM MEMVAR 
*   ENDIF 
   IF !SEEK(m.vir, 'tmpr')
    INSERT INTO tmpr FROM MEMVAR 
   ENDIF 

  ENDSCAN 
  USE 
  
  SELECT profs
  FOR m.i=1 TO 3
   m.tip = m.i
   SET RELATION TO STR(m.tip,1)+profil INTO tmpr
   SCAN 
    m.cod = INT(VAL(profil))
    m.vir = STR(m.tip,1)+STR(m.cod,3)
    m.v = IIF(!EMPTY(tmpr.prv), .t., .f.)
    INSERT INTO &licalias FROM MEMVAR 
   ENDSCAN 
   SET RELATION OFF INTO tmpr
  ENDFOR 
  USE IN &licalias
  USE IN tmpr
  
  SELECT aisoms

 ENDSCAN 
 WAIT CLEAR 
 USE 
RETURN 