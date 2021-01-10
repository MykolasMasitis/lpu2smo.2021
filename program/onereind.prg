FUNCTION OneReind(lcPath)

tc_mcod = SUBSTR(lcPath, RAT('\', lcPath)+1)

IF OpenFile("&lcPath\People", "People", "excl") == 0
 SELECT People 
 DELETE TAG ALL 
 SET FULLPATH OFF 
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON sn_pol TAG sn_pol
 INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
 INDEX on dr TAG dr
 INDEX ON s_all TAG s_all 
 SET FULLPATH OFF 
 USE IN people 
ENDIF

IF OpenFile(lcPath+'\Talon', 'Talon', 'excl') == 0
 SELECT Talon
 DELETE TAG ALL 
 SET FULLPATH OFF 
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON c_i TAG c_i
 INDEX ON sn_pol TAG sn_pol
 INDEX ON otd TAG otd
 INDEX ON pcod TAG pcod
 INDEX ON ds TAG ds
 INDEX ON d_u TAG d_u
 INDEX ON cod TAG cod
 INDEX ON profil TAG profil
 INDEX ON sn_pol+STR(cod,6)+DTOS(d_u) TAG ExpTag
 INDEX ON sn_pol+otd+ds+PADL(cod,6,'0')+DTOC(d_u) TAG unik 
 INDEX ON tip TAG tip
 INDEX ON s_all TAG s_all
 SET FULLPATH OFF 
 USE IN talon 
ENDIF

IF OpenFile(lcPath+'\Otdel', "Otdel", "excl") == 0
 SELECT Otdel
 DELETE TAG ALL 
 SET FULLPATH OFF 
 INDEX ON iotd TAG iotd
 SET FULLPATH OFF 
 USE IN Otdel
ENDIF

IF OpenFile(lcPath+'\Doctor', "Doctor", "excl") == 0
 SELECT Doctor
 DELETE TAG ALL 
 SET FULLPATH OFF 
 INDEX ON pcod TAG pcod
 SET FULLPATH OFF 
 USE IN Doctor
ENDIF

IF fso.FileExists(lcPath+'\ho'+m.qcod)
 IF OpenFile(lcPath+'\ho'+m.qcod, "ho", "excl") == 0
  SELECT ho
  DELETE TAG ALL 
  SET FULLPATH OFF 
  INDEX on sn_pol+c_i+PADL(cod,6,'0') TAG unik
  INDEX on sn_pol+c_i TAG snp_ci
*  INDEX ON pcod TAG pcod
  SET FULLPATH OFF 
  USE IN ho
 ENDIF
ENDIF 

IF fso.FileExists(lcPath+'\onk_sl'+m.qcod+'.dbf')
 IF OpenFile(lcPath+'\ONK_SL'+m.qcod, "onk", "excl") == 0
  SELECT onk
  DELETE TAG ALL 
  SET FULLPATH OFF 
  IF FIELD('recid_s')=UPPER('recid_s')
   INDEX on recid_s TAG recid_s
  ENDIF 
  INDEX on recid TAG recid
  *INDEX ON sn_pol TAG sn_pol
  *INDEX ON c_i TAG c_i
  SET FULLPATH OFF 
  USE IN onk
 ENDIF
ENDIF 

IF fso.FileExists(lcPath+'\onk_usl'+m.qcod+'.dbf')
 IF OpenFile(lcPath+'\ONK_USL'+m.qcod, "onk", "excl") == 0
  SELECT onk
  DELETE TAG ALL 
  SET FULLPATH OFF 
  IF FIELD('recid_sl')=UPPER('recid_sl')
   INDEX on recid_sl TAG recid_s
  ENDIF 
  INDEX on recid TAG recid
  *INDEX ON sn_pol TAG sn_pol
  *INDEX ON c_i TAG c_i
  SET FULLPATH OFF 
  USE IN onk
 ENDIF
ENDIF 

IF fso.FileExists(lcPath+'\onk_ls'+m.qcod+'.dbf')
 IF OpenFile(lcPath+'\ONK_LS'+m.qcod, "onk", "excl") == 0
  SELECT onk
  DELETE TAG ALL 
  SET FULLPATH OFF 
  INDEX on recid_usl TAG recid_s
  SET FULLPATH OFF 
  USE IN onk
 ENDIF
ENDIF 

IF fso.FileExists(lcPath+'\onk_diag'+m.qcod+'.dbf')
 IF OpenFile(lcPath+'\ONK_DIAG'+m.qcod, "onk", "excl") == 0
  SELECT onk
  DELETE TAG ALL 
  SET FULLPATH OFF 
  IF FIELD('recid_sl')=UPPER('recid_sl')
   INDEX on recid_sl TAG recid
  ENDIF 
  SET FULLPATH OFF 
  USE IN onk
 ENDIF
ENDIF 

IF fso.FileExists(lcPath+'\cv_ls'+m.qcod+'.dbf')
 IF OpenFile(lcPath+'\cv_ls'+m.qcod, "onk", "excl") == 0
  SELECT onk
  DELETE TAG ALL 
  SET FULLPATH OFF 
  INDEX on rid TAG rid 
  INDEX on recid_s TAG recid_s
  SET FULLPATH OFF 
  USE IN onk
 ENDIF
ENDIF 

IF fso.FileExists(lcPath+'\onk_cons'+m.qcod+'.dbf')
 IF OpenFile(lcPath+'\ONK_CONS'+m.qcod, "onk", "excl") == 0
  SELECT onk
  DELETE TAG ALL 
  SET FULLPATH OFF 
  INDEX on recid_s TAG recid
  SET FULLPATH OFF 
  USE IN onk
 ENDIF
ENDIF 

IF !fso.FileExists(lcPath+'\e'+m.tc_mcod+'.dbf')
 ErrFile = lcPath + '\e' + tc_mcod
 CREATE TABLE (ErrFile) (f c(1), c_err c(3), rid i)
 USE 
ENDIF 

IF OpenFile(lcPath+'\e'+m.tc_mcod, "Error", "excl") == 0
 SELECT Error
 DELETE TAG ALL 
 SET FULLPATH OFF 
 INDEX FOR UPPER(f)='R' ON rid TAG rrid
 INDEX FOR UPPER(f)='S' ON rid TAG rid
 SET FULLPATH OFF 
 USE IN error
ENDIF

emfile = lcPath+'\m'+m.tc_mcod
IF fso.FileExists(emfile+'.dbf')
 IF OpenFile(emfile, "Error", "excl") == 0
  SELECT Error
  DELETE TAG ALL 
  IF FIELD('rid')!='RID'
   ALTER TABLE error add COLUMN rid i AUTOINC 
  ENDIF 
  INDEX ON rid TAG rid 
  INDEX ON RecId TAG recid
  INDEX ON PADL(recid,6,'0')+et+docexp+reason TAG id_et
  INDEX ON PADL(recid,6,'0')+et+docexp+reason+LEFT(err_mee,2) TAG unik
  INDEX ON PADL(recid,6,'0')+et TAG uniket
  SET FULLPATH OFF 
  SET FULLPATH OFF 
  USE IN error
 ENDIF
ENDIF 

IF !fso.FileExists(lcPath+'\ExpSelected.dbf')
 CREATE TABLE &lcPath\expselected (recid i, sn_pol c(25)) 
 USE 
ENDIF 

IF OpenFile(lcPath+'\ExpSelected', "ExpSel", "excl") == 0
 SELECT ExpSel
 DELETE TAG ALL 
 INDEX ON RecID TAG RecID
 INDEX ON sn_pol TAG sn_pol
 USE IN ExpSel
ENDIF

RETURN 
