FUNCTION SetEnv(para1)
 PUBLIC  m.qcod, m.fso
 m.qcod = para1
 
 SET SAFETY OFF
 fso  = CREATEOBJECT('Scripting.FileSystemObject')
 SET PROCEDURE TO Utils.prg
 
RETURN 

FUNCTION _BasReind

m.t_start = SECONDS()

Local Parallel as Parallel of ParallelFox.vcx

m.IsPparallel = .T.
TRY 
 Parallel = NewObject("Parallel", "ParallelFox.vcx")
CATCH 
 m.IsPparallel = .F.
ENDTRY 

IF m.IsPparallel = .T.
 Parallel.StartWorkers(FullPath("lpu2smo.exe"),,.t.)
 Parallel.Do("SetEnv","BasReind.prg",.T., m.qcod)
ENDIF 

WAIT "Переиндексация aisoms.dbf..." WINDOW NOWAIT 
IF OpenFile(pBase+'\'+gcPeriod+'\AisOms', 'AisOms', 'excl') == 0
 SELECT AisOms
 DELETE TAG ALL 
 INDEX ON cmessage TAG cmessage
 INDEX ON mcod TAG mcod
 INDEX ON lpuid TAG lpuid
 INDEX ON TTOC(sent,1)      TAG sent 
 INDEX ON TTOC(recieved,1)  TAG recieved
 INDEX ON TTOC(processed,1) TAG processed
 INDEX ON paz TAG paz
 INDEX ON s_pred TAG s_pred
 INDEX ON sum_flk TAG sum_flk
 INDEX ON usr TAG usr
 INDEX ON e_mee TAG e_mee 
 INDEX ON e_ekmp TAG e_ekmp
 INDEX ON tpn TAG tpn
 IF FIELD('soapsts')=UPPER('soapsts')
  INDEX on soapsts TAG soapsts
 ENDIF 
 USE 
ENDIF 
WAIT CLEAR 

tn_result = 0
tn_result = tn_result + OpenFile(pBase+'\'+gcPeriod+'\AisOms', 'AisOms', 'shar', 'mcod')
IF tn_result > 0
 RETURN 
ENDIF 

SELECT AisOms
SCAN  
 WAIT mcod WINDOW NOWAIT 
 lcPath = pbase+'\'+m.gcperiod+'\'+mcod
 IF fso.FileExists(lcPath+'\People.dbf') AND ;
    fso.FileExists(lcPath+'\Talon.dbf') AND ;
    fso.FileExists(lcPath+'\Otdel.dbf')
  IF m.IsPparallel = .T.
   Parallel.Do("OneReind","OneReind.prg",, lcPath)
  ELSE 
   =OneReind(lcPath)
  ENDIF 
 ENDIF 
 WAIT CLEAR 
ENDSCAN 
USE 

IF m.IsPparallel = .T.
 Parallel.Wait()
ENDIF 

m.t_end = SECONDS()

m.t_last = (m.t_end - m.t_start)

MESSAGEBOX('Время: '+TRANSFORM(m.t_last, '999999.99'),0+64,'')

RETURN 

FUNCTION BasReind(period)

 DO CASE 
  CASE m.ffoms = 77011
   m.LoadPeriod={01.02.2012}+5*365
  CASE m.ffoms = 77002
   m.LoadPeriod={15.02.2012}+5*365 && 13.02.2017
  CASE m.ffoms = 77008
   m.LoadPeriod={17.02.2012}+5*365 && 15.02.2017
  CASE m.ffoms = 77013
   m.LoadPeriod={15.03.2012}+5*365 && 14.03.2017
  CASE m.ffoms = 77012
   m.LoadPeriod={17.03.2012}+5*365 && 16.03.2017
  OTHERWISE 
   m.LoadPeriod={17.03.2012}+5*365 && 16.03.2017

 ENDCASE 

IF DATE()>m.LoadPeriod
* =ChkDirsBrief()
ENDIF 

WAIT "Переиндексация aisoms.dbf..." WINDOW NOWAIT 
IF OpenFile(pBase+'\'+period+'\AisOms', 'AisOms', 'excl') == 0
 SELECT AisOms
 DELETE TAG ALL 
 INDEX ON cmessage TAG cmessage
 INDEX ON mcod TAG mcod
 INDEX ON lpuid TAG lpuid
 INDEX ON TTOC(sent,1)      TAG sent 
 INDEX ON TTOC(recieved,1)  TAG recieved
 INDEX ON TTOC(processed,1) TAG processed
 INDEX ON paz TAG paz
 INDEX ON s_pred TAG s_pred
 INDEX ON sum_flk TAG sum_flk
 INDEX ON usr TAG usr
 INDEX ON e_mee TAG e_mee 
 INDEX ON e_ekmp TAG e_ekmp
 INDEX ON tpn TAG tpn
 IF FIELD('soapsts')=UPPER('soapsts')
  INDEX on soapsts TAG soapsts
 ENDIF 
 USE 
ENDIF 
WAIT CLEAR 

IF fso.FileExists(pBase+'\'+period+'\people.dbf')
 WAIT "Переиндексация people.dbf..." WINDOW NOWAIT 
 =BigPeopleReindex(pBase+'\'+period+'\people')
 WAIT CLEAR 
ENDIF 

IF fso.FileExists(pBase+'\'+period+'\dsp.dbf')
 WAIT "Переиндексация dsp.dbf..." WINDOW NOWAIT 
 IF OpenFile(pBase+'\'+period+'\dsp', 'dsp', 'excl') == 0
  SELECT dsp
  DELETE TAG ALL 
  INDEX ON period+mcod+PADL(recid,6,'0') TAG uniqq
  INDEX on sn_pol+PADL(tip,1,'0') TAG exptag
  INDEX on sn_pol+PADL(cod,6,'0') TAG un_tag
  INDEX on c_i TAG c_i 
  USE 
 ENDIF 
 WAIT CLEAR 
ENDIF 

pr4file = pBase+'\'+period+'\pr4'
IF fso.FileExists(pr4file+'.dbf')
 WAIT "Переиндексация pr4.dbf..." WINDOW NOWAIT 
 IF OpenFile(pr4file, 'pr4', 'excl')<=0
  SELECT pr4
  DELETE TAG ALL 
  INDEX ON lpuid TAG lpuid
  INDEX on mcod TAG mcod
  USE 
 ENDIF 
 WAIT CLEAR 
ENDIF 

IF fso.FileExists(pBase+'\'+period+'\talon.dbf')
 WAIT "Переиндексация talon.dbf..." WINDOW NOWAIT 
 =BigTalonReindex(pBase+'\'+period+'\talon')
 WAIT CLEAR 
ENDIF  

IF fso.FileExists(pBase+'\'+period+'\otdel.dbf')
 WAIT "Переиндексация otdel.dbf..." WINDOW NOWAIT 
 =BigOtdelReindex(pBase+'\'+period+'\otdel')
 WAIT CLEAR 
ENDIF 

IF fso.FileExists(pBase+'\'+period+'\doctor.dbf')
 WAIT "Переиндексация doctor.dbf..." WINDOW NOWAIT 
 =BigDoctorReindex(pBase+'\'+period+'\otdel')
 WAIT CLEAR 
ENDIF 

errsv  = pBase+'\'+period+'\e'+m.period
IF fso.FileExists(errsv+'.dbf')
WAIT "Переиндексация e"+m.period+".dbf..." WINDOW NOWAIT 
IF OpenFile(errsv, 'errsv', 'excl') == 0
 SELECT errsv
 DELETE TAG ALL 
 INDEX FOR UPPER(f)='R' ON rid TAG rrid
 INDEX FOR UPPER(f)='S' ON rid TAG rid
 USE 
ENDIF 
WAIT CLEAR 
ENDIF 

IF fso.FileExists(pBase+'\'+period+'\mee.dbf')
WAIT "Переиндексация mee.dbf..." WINDOW NOWAIT 
IF OpenFile(pBase+'\'+period+'\mee', 'mee', 'excl') == 0
 SELECT mee
 DELETE TAG ALL 
 INDEX ON rid TAG rid
 USE 
ENDIF 
WAIT CLEAR 
ENDIF 

IF fso.FileExists(pbase+'\'+period+'\expdetails.dbf')
 WAIT "Переиндексация expdetails.dbf..." WINDOW NOWAIT 
 IF OpenFile(pbase+'\'+period+'\expdetails', 'expdetails', 'excl') == 0
  SELECT expdetails
  DELETE TAG ALL 
  INDEX ON period+mcod+et TAG ikey
  INDEX ON mcod TAG mcod 
  INDEX ON et TAG et
  USE
 ENDIF 
 WAIT CLEAR 
ENDIF 

IF fso.FileExists(pmee+'\ssacts\ssacts.dbf')
WAIT "Переиндексация ssacts.dbf..." WINDOW NOWAIT 
IF OpenFile(pmee+'\ssacts\ssacts', 'ssacts', 'excl') == 0
 SELECT ssacts
 DELETE TAG ALL 
 INDEX ON recid TAG recid CANDIDATE 
 INDEX FOR qr ON recid TAG qrrecid 
 INDEX ON period TAG period
 INDEX ON e_period TAG e_period
 INDEX ON mcod TAG mcod 
 INDEX ON sn_pol TAG sn_pol
 INDEX ON actdate TAG actdate
 INDEX ON PADR(ALLTRIM(fam)+' '+LEFT(im,1)+LEFT(ot,1),28) TAG fio 
 USE 
ENDIF 
WAIT CLEAR 
ENDIF 

IF fso.FileExists(pmee+'\ssacts\svacts.dbf')
WAIT "Переиндексация svacts.dbf..." WINDOW NOWAIT 
IF OpenFile(pmee+'\svacts\svacts', 'svacts', 'excl') == 0
 SELECT svacts
 DELETE TAG ALL 
 INDEX ON recid TAG recid 
 INDEX FOR qr ON recid TAG qrrecid 
 INDEX ON period TAG period
 INDEX ON e_period TAG e_period
 INDEX ON mcod TAG mcod 
 INDEX ON actdate TAG actdate
 INDEX ON period+e_period+mcod+STR(codexp,1)+docexp TAG unik
 INDEX on status TAG status 
 USE 
ENDIF 
WAIT CLEAR 
ENDIF 

tn_result = 0
tn_result = tn_result + OpenFile(pBase+'\'+period+'\AisOms', 'AisOms', 'shar', 'mcod')
IF tn_result > 0
 RETURN 
ENDIF 

SELECT AisOms
SCAN  
 WAIT mcod WINDOW NOWAIT 
 lcPath = pbase+'\'+m.period+'\'+mcod
 IF fso.FileExists(lcPath+'\People.dbf') AND ;
    fso.FileExists(lcPath+'\Talon.dbf') AND ;
    fso.FileExists(lcPath+'\Otdel.dbf')
  =OneReind(lcPath)
 ENDIF 
 WAIT CLEAR 
ENDSCAN 
USE 

IF fso.FolderExists(pMee)
 IF fso.FolderExists(pMee+'\REQUESTS')
  IF fso.FileExists(pMee+'\REQUESTS\Catalog.dbf')
   IF OpenFile(pMee+'\REQUESTS\Catalog', 'cat', 'shar')>0
    IF USED('cat')
     USE IN cat
    ENDIF 
   ELSE 
    SELECT cat
    SCAN 
     m.recid = recid
     m.f_name = pMee+'\REQUESTS\'+PADL(m.recid,6,'0')
     IF fso.FileExists(m.f_name+'.dbf')
      IF OpenFile(m.f_name, 'fn', 'excl')>0
       IF USED('fn')
        USE IN fn
       ENDIF 
       SELECT cat
      ELSE
       SELECT fn 
       DELETE TAG all
       INDEX on recid TAG recid
       INDEX on sn_pol TAG sn_pol
       USE 
       SELECT cat
      ENDIF 
     ENDIF 
    ENDSCAN 
    USE IN cat 
   ENDIF 
  ENDIF 
 ENDIF 
ENDIF 

RETURN 

FUNCTION BigPeopleReindex(para1)
 LOCAL lPath
 lPath = para1
 IF fso.FileExists(lPath+'.dbf')
  IF OpenFile(lPath, 'people', 'excl') == 0
   SELECT people
   DELETE TAG ALL 
   INDEX ON RecId TAG recid CANDIDATE 
   INDEX ON sn_pol TAG sn_pol
   INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
   INDEX ON dr TAG dr
   INDEX ON nlpu TAG nlpu
   INDEX on s_all TAG s_all
   USE 
  ENDIF 
 ENDIF 
RETURN 

FUNCTION BigTalonReindex(para1)
 LOCAL lPath
 lPath = para1
 IF fso.FileExists(lPath+'.dbf')
  IF OpenFile(lPath, 'talon', 'excl') == 0
   SELECT talon
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
   USE 
  ENDIF 
 ENDIF 
RETURN 

FUNCTION BigOtdelReindex(para1)
 LOCAL lPath
 lPath = para1
 IF fso.FileExists(lPath+'.dbf')
  IF OpenFile(lPath, 'otdel', 'excl') == 0
   SELECT otdel
   DELETE TAG ALL 
   INDEX ON iotd TAG iotd
   INDEX ON mcod+' '+iotd TAG unkey
   USE 
  ENDIF 
 ENDIF 
RETURN 

FUNCTION BigDoctorReindex(para1)
 LOCAL lPath
 lPath = para1
 IF fso.FileExists(lPath+'.dbf')
  IF OpenFile(lPath, 'doctor', 'excl') == 0
   SELECT doctor
   DELETE TAG ALL 
   USE 
  ENDIF 
 ENDIF 
RETURN 