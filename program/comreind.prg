PROCEDURE ComReind

WAIT "œÂÂËÌ‰ÂÍ‡ˆËˇ COMMON..." WINDOW NOWAIT 

m.lWasUsed = .F.
IF USED('Users')
 m.lWasUsed = .T.
 SELECT Users
 COUNT FOR ISRLOCKED() TO m.nLocked
 IF m.nLocked > 0
  SELECT name FROM Users WHERE !ISRLOCKED() INTO CURSOR wlck
  SELECT wlck
  INDEX on name TAG name 
  SET ORDER TO name 
 ENDIF 
 USE IN users
ENDIF 
IF OpenFile(pcommon+'\Users', 'users', 'excl') <= 0
 SELECT Users 
 INDEX on name TAG name 
 USE 
ENDIF 
IF m.lWasUsed=.T.
 =OpenFile(pCommon+'\Users', 'Users', 'shar', 'name')
 IF USED('wlck')
  SELECT Users
  SET RELATION TO name INTO wlck
  SCAN 
   IF EMPTY(wlck.name)
    RLOCK()
   ENDIF 
  ENDSCAN 
  SET RELATION OFF INTO wlck
  USE IN wlck
 ENDIF 
ENDIF 

IF !fso.FileExists(pCommon+'\Users.cdx')
 IF OpenFile(pcommon+'\Users', 'users', 'excl') <= 0
  SELECT Users 
  INDEX on name TAG name 
  USE 
 ENDIF 
ENDIF 

IF OpenFile(pCommon+'\explist', "explist", "excl") == 0
 SELECT explist
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
* INDEX ON cod FOR comidx() TAG cod
 INDEX ON cod TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pCommon+'\lpuskp', "lpuskp", "excl") == 0
 SELECT lpuskp
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpuid TAG lpuid
 INDEX on mcod TAG mcod 
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pCommon+'\dc_du', "dcdu", "excl") == 0
 SELECT dcdu
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpu_id FOR dc TAG dc 
 INDEX ON lpu_id FOR du TAG du
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pCommon+'\pnyear', "pnyear", "excl") == 0
 SELECT pnyear
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 DELETE TAG ALL 
 INDEX ON period TAG period
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pCommon+'\pnorm', "pnorm", "excl") == 0
 SELECT pnorm
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 DELETE TAG ALL 
* INDEX ON period FOR comidx() TAG period
 INDEX ON period TAG period
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pCommon+'\pnorms', "pnorm", "excl") == 0
 SELECT pnorm
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 DELETE TAG ALL 
* INDEX ON period FOR comidx() TAG period
 INDEX ON period TAG period
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\prv002xx', "prv002", "excl") == 0
 SELECT prv002
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
* INDEX ON profil FOR comidx() TAG profil
 INDEX ON profil TAG profil
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pCommon+'\emails', "emails", "excl") == 0
 SELECT emails
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpu_id TAG lpu_id
 INDEX ON mcod TAG mcod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pCommon+'\dsdisp', "dsdisp", "excl") == 0
 SELECT dsdisp
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpu_id TAG lpu_id
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\smo', "smo", "excl") == 0
 SELECT smo
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON code TAG code
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\lpudogs', 'lpudogs', "excl") == 0
 SELECT lpudogs
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpu_id TAG lpu_id
 INDEX ON mcod TAG mcod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\dspcodes', "dspcodes", "excl") == 0
 SELECT dspcodes
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cod TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\n_kd', "n_kd", "excl") == 0
 SELECT n_kd
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cod TAG cod 
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\reasons', "reasons", "excl") == 0
 SELECT reasons
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON code TAG code
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\exptip', "exptip", "excl") == 0
 SELECT exptip
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON code TAG code
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\mee2mgf', "mee2mgf", "excl") == 0
 SELECT mee2mgf
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON my_et TAG my_et
 INDEX ON mgf_et TAG mgf_et
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\usvmpxx', "usvmp", "excl") == 0
 SELECT usvmp
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cod TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\hopff_xx', "hopff", "excl") == 0
 SELECT hopff
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON codho TAG codho
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\et', "et", "excl") == 0
 SELECT et
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON code TAG code
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\errsmee', "errsmee", "excl") == 0
 SELECT errsmee
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON er_c TAG er_c
 INDEX on osn230 TAG osn230
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\rsltishod', "rr", "excl") == 0
 SELECT rr
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON STR(rslt,3)+STR(ishod,3) TAG unik
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\pervpr', "ppr", "excl") == 0
 SELECT ppr
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cod TAG cod 
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

WAIT CLEAR 

WAIT "œÂÂËÌ‰ÂÍ‡ˆËˇ NSI..." WINDOW NOWAIT 

IF OpenFile(pbase+'\'+gcperiod+'\nsi\'+'errsmee', "errsmee", "excl") == 0
 SELECT errsmee
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON er_c TAG er_c
 INDEX on osn230 TAG osn230
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\nsi\pilot', "pilot", "excl") == 0
 SELECT pilot
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpu_id TAG lpu_id
 INDEX ON mcod TAG mcod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilots', "pilot", "excl") == 0
 SELECT pilot
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpu_id TAG lpu_id
 INDEX ON mcod TAG mcod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\noth', "noth", "excl") == 0
 SELECT noth
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cod TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\admokrxx', "admokr", "excl") == 0
 SELECT admokr
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cokr TAG cokr
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\CodKu', "CodKU", "excl") == 0
 SELECT codku
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cod TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\CodOtd', "codotd", "excl") == 0
 SELECT codotd
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cod FOR kl='#' TAG ncod && “ÓÎ¸ÍÓ ‚ Ú‡ÍËı ÓÚ‰ÂÎÂÌËˇı ˝Ú‡ ÛÒÎÛ¯‡
 INDEX ON otd FOR kl='y' TAG notd && “ÓÎ¸ÍÓ Ú‡ÍËÂ ÛÒÎÛ„Ë ‚ ˝ÚÓÏ ÓÚ‰ÂÎÂÌËË
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\codwdr', "codwdr", "excl") == 0
 SELECT codwdr
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cod TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\hopff', "hopff", "excl") == 0
 SELECT hopff
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON codho TAG codho
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\isv012', "isv012", "excl") == 0
 SELECT isv012
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON ishod TAG ishod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\kdolgxx', "kdolg", "excl") == 0
 SELECT kdolg
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON prvd TAG prvd
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\kpresl', "tips", "excl") == 0
 SELECT tips
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON tip TAG tip 
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF m.tdat1<{01.05.2014}
IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\kspec', "kspec", "excl") == 0
 SELECT kspec
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON prvs_foms TAG prvs
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF
ENDIF 

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\lputpn', "lputpn", "excl") == 0
 SELECT lputpn
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpu_id TAG lpu_id
 INDEX ON fil_id TAG fil_id 
 INDEX ON mcod TAG mcod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\horlpu', "horlpu", "excl") == 0
 SELECT horlpu
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpu_id TAG lpu_id
 INDEX ON fil_id TAG fil_id 
 INDEX ON mcod TAG mcod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\horlpus', "horlpu", "excl") == 0
 SELECT horlpu
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpu_id TAG lpu_id
 INDEX ON fil_id TAG fil_id 
 INDEX ON mcod TAG mcod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\mkb10', "mkb10", "excl") == 0
 SELECT mkb10
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON ds TAG ds
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

*IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\mo_vmp', "movmp", "excl") == 0
* SELECT movmp
* SET FULLPATH OFF 
* WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
* INDEX ON lpu_id TAG lpu_id
* SET FULLPATH OFF 
* USE
* WAIT CLEAR 
*ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\modpac', "modpac", "excl") == 0
 SELECT modpac
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON codmod TAG codmod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\ms_mkb', "msmkb", "excl") == 0
 SELECT msmkb
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON ds TAG ds
 INDEX ON COD TAG COD
 INDEX ON STR(cod,6)+' '+ds TAG ds_ms
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\nocodr', "nocodr", "excl") == 0
 SELECT nocodr
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON LEFT(Ds,1) FOR Maska=1 TAG Ds1
 INDEX ON LEFT(Ds,2) FOR Maska=2 TAG Ds2
 INDEX ON LEFT(Ds,3) FOR Maska=3 TAG Ds3
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\osoerzxx', "OsoERZ", "excl") == 0
 SELECT OsoERZ
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON Ans_r TAG Ans_r
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\osoree', "osoree", "excl") == 0
 SELECT osoree
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON d_type TAG d_type
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\ososch', "ososch", "excl") == 0
 SELECT ososch
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON d_type TAG d_type
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

*IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\polic_dp', "PolisDP", "excl") == 0
* SELECT PolisDP
* SET FULLPATH OFF 
* WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
* INDEX ON sn_pol TAG sn_pol
* SET FULLPATH OFF 
* USE
* WAIT CLEAR 
*ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\polic_h', "PolisH", "excl") == 0
 SELECT PolisH
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON sn_pol TAG sn_pol
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\profot', "profot", "excl") == 0
 SELECT profot
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON otd TAG otd
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\profus', "profus", "excl") == 0
 SELECT profus
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cod TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\rsv009', "rsv009", "excl") == 0
 SELECT rsv009
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON rslt TAG rslt
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\sookodxx', "sookod", "excl") == 0
 SELECT sookod
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON er_c TAG er_c
 INDEX ON osn230 TAG osn230
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\SovmNO', "SovmNO", "excl") == 0
 SELECT SovmNO
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX FOR Kl='#' ON Cod TAG ncod 
 INDEX FOR Kl='y' ON Cod TAG scod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

*IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\spi_lpu_dd', "spi", "excl") == 0
* SELECT spi
* SET FULLPATH OFF 
* WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
* INDEX ON lpu_id TAG lpu_id
* SET FULLPATH OFF 
* USE
* WAIT CLEAR 
*ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\spraboxx', "sprabo", "excl") == 0
 SELECT sprabo
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON abn_name TAG abn_name
 INDEX ON object_id TAG lpu_id
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\sprlpuxx', "sprlpu", "excl") == 0
 SELECT sprlpu
 IF VARTYPE(fil_id) != 'N'
  ALTER TABLE sprlpu ADD COLUMN fil_id n(6)
  REPLACE ALL fil_id WITH lpu_id
 ENDIF 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
* INDEX ON mcod FOR lpu_id=fil_id and du_in>=m.t_dat1 TAG mcod
* INDEX FOR lpu_id=fil_id and du_in>=m.tdat1 ON lpu_id TAG lpu_id
* INDEX FOR lpu_id=fil_id and du_in>=m.tdat1 ON mcod TAG mcod
* INDEX FOR lpu_id=fil_id and du_in>=m.tdat1 ON cokr TAG cokr
 INDEX ON lpu_id TAG lpu_id
 INDEX ON fil_id TAG fil_id
 INDEX ON mcod TAG mcod
 INDEX ON cokr TAG cokr
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\spv015', "spv015", "excl") == 0
 SELECT spv015
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON code TAG code
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\street', "street", "excl") == 0
 SELECT street
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON ul TAG ul
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\tarifn', "tarifn", "excl") == 0
 SELECT tarifn
* m.nIsTrouble = 0
* COUNT FOR ;
  (INT(cod/1000)=183 AND stkd != ROUND(tarif/7,2)) OR (INT(cod/1000)=183 AND stkdv != ROUND(tarif_v/7,2));
   TO m.nIsTrouble
* IF m.nIsTrouble>0
*  IF MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡Õ¿–”∆≈Õ¿ œ–Œ¡À≈Ã¿ ¬ “¿–»‘≈! »—œ–¿¬»“‹?'+CHR(13)+CHR(10),4+32,'')!=7
*   REPLACE FOR INT(cod/1000)=183 AND stkd != ROUND(tarif/7,2) stkd WITH ROUND(tarif/7,2)
*   REPLACE FOR INT(cod/1000)=183 AND stkdv != ROUND(tarif_v/7,2) stkdv WITH ROUND(tarif_v/7,2)
*  ENDIF 
* ENDIF 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cod TAG cod
 *REPLACE FOR IsVmp(cod) OR IsMes(cod) tpn WITH ''
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\tarion', "tarion", "excl") == 0
 SELECT tarion
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod TAG cod
 INDEX FOR edizm="ÏÎ" ON cod+TRANSFORM(mass_value,"999.99") TAG ML_COD
 INDEX FOR edizm="Ï„" ON cod+TRANSFORM(mass_value,"999.99") TAG MG_COD
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\medicament', "medicament", "excl") == 0
 SELECT medicament
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on dd_sid TAG dd_sid
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\medicament_mfc', "medicament", "excl") == 0
 SELECT medicament
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on dd_id TAG dd_id
 INDEX on n_ru TAG n_ru
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\medpack', "medpack", "excl") == 0
 SELECT medpack
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on r_up TAG r_up
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\sprnco', "sprnco", "excl") == 0
 SELECT sprnco
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on lpu_id TAG lpu_id
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\prvr2', "prvr2", "excl") == 0
 SELECT prvr2
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on profil TAG profil
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\nsif', "nsif", "excl") == 0
 SELECT nsif
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on lpu_id TAG lpu_id
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\nsio', "nsio", "excl") == 0
 SELECT nsio
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on STR(lpu_id,4)+STR(n_str,2) TAG unik
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\gr_plan', "gr", "excl") == 0
 SELECT gr
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\ms_ds_prv', "nnn", "excl") == 0
 SELECT nnn
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on STR(cod,6)+ds TAG unik
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\perv', "pervpr", "excl") == 0
 SELECT pervpr
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pcommon+'\wom', "wom", "excl") == 0
 SELECT wom
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpu_id TAG lpu_id
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\tipgrp', "tipgrp", "excl") == 0
 SELECT tipgrp
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON tipgr TAG tipgr
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\UsrLpu', "UsrLpu", "excl") == 0
 SELECT UsrLpu
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON mcod TAG mcod
 INDEX ON lpu_id TAG lpu_id
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\vidvp', "vidvp", "excl") == 0
 SELECT vidvp
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON codvvp TAG codvvp
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\tipnomes', "tipnomes", "excl") == 0
 SELECT tipnomes
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON STR(cod,6)+' '+Tip TAG vir
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\reeskp', "reeskp", "excl") == 0
 SELECT reeskp
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod TAG cod UNIQUE 
 INDEX ON PADL(cod,6,'0')+codho+ds TAG unik
 INDEX on STR(cod,6)+" "+ds TAG ds_ms 
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onmet_xx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_m TAG cod_m
 INDEX on ds TAG ds
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onnod_xx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_n TAG cod_n
 INDEX on ds TAG ds
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onreasxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_reas TAG cod_reas
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onstadxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_st TAG cod_st
 INDEX on ds TAG ds
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\ontum_xx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_t TAG cod_t
 INDEX on ds TAG ds
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onlechxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_tlech TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onhir_xx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_thir TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onleklxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_tlek_l TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onlekvxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_tlek_v TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onluchxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_tluch TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onprotxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_prot TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onmrf_xx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_mrf TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onmrdsxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_mrf TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onigh_xx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_igh TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onigdsxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_igh TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onmrfrxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on id_r_m TAG id_r_m
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onigrtxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on id_r_i TAG id_r_i
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onconsxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_cons TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onpcelxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_pc TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onnaprxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_vn TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onczabxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_cz TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\ondopkxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod_dkk TAG cod_dkk
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onoplsxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on ds UNIQUE tag ds
 INDEX on n020 UNIQUE TAG regnum 
 INDEX on PADL(cod_ms,6,'0')+STR(n020,4)+ds tag unik 
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onlpshxx', "onk", "excl") == 0
 SELECT onk
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON code_sh TAG code_sh 
 INDEX ON code_sh+id_lekp TAG unik
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\msext', "msext", "excl") == 0
 SELECT msext
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod tag cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\sprved', "sprved", "excl") == 0
 SELECT sprved
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on lpu_id TAG lpu_id
 INDEX on mcod TAG mcod 
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\tarion', "tarion", "excl") == 0
 SELECT tarion
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cod TAG cod
 INDEX ON cod + TRANSFORM(mass_value,'999.99') TAG ml_cod FOR edizm = 'ÏÎ'
 INDEX ON cod + TRANSFORM(mass_value,'999.99') TAG mg_cod FOR edizm = 'Ï„'
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\codprv', "cprv", "excl") == 0
 SELECT cprv
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod TAG cod 
 INDEX on STR(cod,6)+STR(code,3) TAG var
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi\kms.dbf')
IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\kms', "kms", "excl") == 0
 SELECT kms
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON vs TAG vs
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF
WAIT CLEAR 
ENDIF 

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\pnorm_iskl', "pn", "excl") == 0
 SELECT pn
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON mcod TAG mcod
 INDEX ON lpu_id TAG lpu_id
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF
WAIT CLEAR 

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\ns36', "ns", "excl") == 0
 SELECT ns
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cod TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF
WAIT CLEAR 

WAIT "œ≈–≈»Õƒ≈ —¿÷»ﬂ Ã››..." WINDOW NOWAIT 
IF OpenFile(pmee+'\ssacts\ssacts', 'ssacts', 'excl') == 0
 SELECT ssacts
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX FOR qr ON recid TAG qrrecid 
 INDEX ON recid TAG recid 
 INDEX ON period TAG period
 INDEX ON mcod TAG mcod 
 INDEX ON sn_pol TAG sn_pol
 INDEX ON actdate TAG actdate
 INDEX ON PADR(ALLTRIM(fam)+' '+LEFT(im,1)+LEFT(ot,1),28) TAG fio 
 INDEX on status TAG status
 USE
ENDIF 
WAIT CLEAR 

IF OpenFile(pmee+'\ssacts\moves', 'moves', 'excl') == 0
 SELECT moves
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON recid TAG recid 
 INDEX ON actid TAG actid 
 USE
 WAIT CLEAR 
ENDIF 

IF OpenFile(pmee+'\svacts\svacts', 'svacts', 'excl') == 0
 SELECT svacts
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON recid TAG recid 
 INDEX FOR qr ON recid TAG qrrecid 
 INDEX ON period TAG period
 INDEX ON e_period TAG e_period
 INDEX ON mcod TAG mcod 
 INDEX ON actdate TAG actdate
* INDEX ON period+et TAG unik 
 INDEX ON period+e_period+mcod+STR(codexp,1)+docexp TAG unik
 INDEX on status TAG status
 USE
 WAIT CLEAR 
ENDIF 

IF OpenFile(pmee+'\svacts\moves', 'moves', 'excl') == 0
 SELECT moves
 DELETE TAG ALL 
 SET FULLPATH OFF 
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON recid TAG recid 
 INDEX ON actid TAG actid 
 USE
 WAIT CLEAR 
ENDIF 

IF OpenFile(pmee+'\rss\rss', 'rss', 'excl') == 0
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 SELECT rss
 DELETE TAG ALL 
 SET FULLPATH OFF 
 INDEX on smoexp TAG smoexp 
 INDEX ON mcod+DTOS(d_u) TAG unik
 INDEX ON recid TAG recid 
 INDEX ON e_period TAG e_period
 INDEX ON lpu_id TAG lpu_id
 INDEX ON mcod TAG mcod 
 USE 
 WAIT CLEAR 
ENDIF 

IF OpenFile(pmee+'\requests\catalog', 'catalog', 'excl') == 0
 WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 SELECT catalog
 DELETE TAG ALL 
 SET FULLPATH OFF 
 INDEX ON recid TAG recid 
 INDEX on mcod TAG mcod 
 INDEX on lpu_id TAG lpu_id
 INDEX on period TAG period
 INDEX on e_period TAG e_period
 INDEX on smoexp TAG smoexp
 INDEX ON smoexp+mcod+period TAG s_unik
 INDEX ON smoexp+mcod+period+et TAG unik
 SCAN 
  m.iid = recid
  m.cid = PADL(m.iid,6,'0')
  IF fso.FileExists(pmee+'\requests\'+m.cid+'.dbf')
   IF OpenFile(pmee+'\requests\'+m.cid, 'rqst', 'excl')=0
    SELECT rqst
    DELETE TAG ALL 
    INDEX ON sn_pol TAG recid
    USE IN rqst
   ENDIF 
  ELSE 
   IF USED('rqst')
    USE IN rqst
   ENDIF 
  ENDIF 
  SELECT catalog
 ENDSCAN 
 USE IN catalog 
 WAIT CLEAR 
ENDIF 

*=DeleteFile('codexpxx.dbf')
=DeleteFile('dp_0112.dbf')
=DeleteFile('emails.dbf')
=DeleteFile('emails.cdx')
=DeleteFile('h_0112.dbf')
=DeleteFile('im1.dbf')
=DeleteFile('im1.cdx')
=DeleteFile('im2.dbf')
=DeleteFile('im2.cdx')
=DeleteFile('loggfile.dbf')
=DeleteFile('lpu_m.dbf')
=DeleteFile('lpu_m.cdx')
=DeleteFile('nomlpu.dbf')
=DeleteFile('pos_dom.dbf')
=DeleteFile('pos_dom.cdx')
=DeleteFile('prilpuxx.dbf')
=DeleteFile('rsltatxx.dbf')
=DeleteFile('spr_mo.dbf')
=DeleteFile('sprsmo.dbf')
=DeleteFile('sprsmo.cdx')
=DeleteFile('stac_mod.dbf')
=DeleteFile('stac_mod.cdx')
=DeleteFile('stmdr.dbf')
=DeleteFile('stmdr.cdx')
=DeleteFile('tar_s.dbf')
=DeleteFile('tar_s.cdx')
=DeleteFile('tarif.dbf')
=DeleteFile('tarimu48.dbf')
=DeleteFile('tarimu48.cdx')
=DeleteFile('tipabo.dbf')
=DeleteFile('users.dbf')
=DeleteFile('users.cdx')
=DeleteFile('usl_m.dbf')
=DeleteFile('usl_m.cdx')
=DeleteFile('usl_obr.dbf')
=DeleteFile('usl_obr.cdx')
=DeleteFile('usl_pos.dbf')
=DeleteFile('usl_pos.cdx')
*=DeleteFile('usrlpu.dbf')
*=DeleteFile('usrlpu.cdx')
=DeleteFile('volumes.dbf')
=DeleteFile('volumes.cdx')
=DeleteFile('z_cod.dbf')
=DeleteFile('z_dsno.dbf')

FUNCTION DeleteFile(m.FileToDeleteShort)
 m.FileToDelete = pbase+'\'+gcperiod+'\'+'nsi\'+m.FileToDeleteShort
 IF fso.FileExists(m.FileToDelete)
  fso.DeleteFile(m.FileToDelete)
 ENDIF 
RETURN 