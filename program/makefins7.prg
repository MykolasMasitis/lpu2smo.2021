PROCEDURE MakeFinS7

IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ СФОРМИРОВАТЬ ФИН-ФАЙЛ?'+CHR(13)+CHR(10),4+32,'')=7
 RETURN 
ENDIF 

IF OpenFile("&pBase\&gcPeriod\aisoms", "aisoms", "shar", "mcod") > 0
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 RETURN .f.
ENDIF 
IF OpenFile(pbase+'\'+gcperiod+'\nsi\'+"sprlpuxx", "sprlpu", "shar", "mcod") > 0
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 RETURN .f.
ENDIF 
m.lIsLpuTpn=.f.
IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\lputpn.dbf')
 m.lIsLpuTpn = .t.
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\lputpn', "lputpn", "shar", "lpu_id") > 0
  IF USED('lputpn')
   USE IN lputpn
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN .f.
 ENDIF 
ENDIF 
IF OpenFile(pbase+'\'+m.gcperiod+'\pr4', 'pr4', 'shar', 'mcod')>0
 IF USED('lputpn')
  USE IN lputpn
 ENDIF 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('pr4')
  USE IN pr4
 ENDIF 
 RETURN .f.
ENDIF 
IF OpenFile(pbase+'\'+m.gcperiod+'\pr4st', 'pr4st', 'shar', 'mcod')>0
 IF USED('lputpn')
  USE IN lputpn
 ENDIF 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('pr4')
  USE IN pr4
 ENDIF 
 IF USED('pr4st')
  USE IN pr4st
 ENDIF 
 RETURN .f.
ENDIF 

IF OpenFile(pbase+'\'+m.gcperiod+'\formmag02', 'mag', 'shar', 'mcod')>0
 IF USED('mag')
  USE IN mag
 ENDIF 
 IF USED('lputpn')
  USE IN lputpn
 ENDIF 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('pr4')
  USE IN pr4
 ENDIF 
 IF USED('pr4st')
  USE IN pr4st
 ENDIF 
 RETURN .f.
ENDIF 

SELECT AisOms
SET RELATION TO mcod INTO SprLpu
SET RELATION TO mcod INTO pr4 ADDITIVE 
SET RELATION TO mcod INTO pr4st ADDITIVE 
SET RELATION TO mcod INTO mag ADDITIVE 

IF OpenFile(pcommon+'\mee2mgf', 'mee2mgf', 'shar', 'my_et')>0
 IF USED('mee2mgf')
  USE IN mee2mgf
 ENDIF 
ENDIF 

m.llIsMeFiles = .T.
m.me_dir        = m.pOut+'\'+m.gcperiod
IF fso.FileExists(m.me_dir+'\ME'+m.qcod+'.dbf')
 fso.DeleteFile(m.me_dir+'\ME'+m.qcod+'.dbf')
ENDIF 

oMailDir        = fso.GetFolder(m.me_dir)
MailDirName     = oMailDir.Path
oFilesInMailDir = oMailDir.Files
nFilesInMailDir = oFilesInMailDir.Count

CREATE CURSOR me&qcod (lpu_id n(4), period_e c(6), et c(1), s_opl_e n(12,2), s_sank n(12,2))
SELECT me&qcod


IF fso.FileExists(m.me_dir+'\full_me'+m.qcod+'.dbf')
 APPEND FROM &me_dir\full_me&qcod
ELSE 
 IF nFilesInMailDir>0
  FOR EACH oFileInMailDir IN oFilesInMailDir
   m.BFullName = oFileInMailDir.Path
   m.bname     = oFileInMailDir.Name
   IF RIGHT(UPPER(m.bname),3)<>'DBF'
    LOOP 
   ENDIF 
   IF LEFT(UPPER(m.bname),2)<>'ME'
    LOOP 
   ENDIF 
  
   APPEND FROM &BFullName

  ENDFOR 
 ENDIF 
ENDIF 

SCAN 
 m.et = et
 m.et = IIF(SEEK(m.et, 'mee2mgf'), mee2mgf.mgf_et, m.et)
 REPLACE et WITH m.et
ENDSCAN 
 
COPY TO &me_dir\ME&qcod
 
SELECT m.gcperiod as period, period_e as e_period, lpu_id as lpuid,;
	et as et, SUM(s_opl_e) as sexp, 000.00 as stpn, SUM(s_sank) as s_sank;
	FROM me&qcod WHERE s_opl_e<>0 OR s_sank<>0 ;
	GROUP BY lpu_id, period_e, et INTO TABLE &me_dir\svmee
SELECT svmee 
INDEX on lpuid TAG lpuid 
USE IN svmee 
USE IN me&qcod
 	
RELEASE m.me_dir, oMailDir, MailDirName, oFilesInMailDir, nFilesInMailDir

IF m.llIsMeFiles
 IF OpenFile(pout+'\'+gcperiod+'\svmee', 'svmee', 'shar', 'lpuid')>0
  IF USED('svmee')
   USE IN svmee
  ENDIF 
  USE IN mee2mgf
  RETURN 
 ENDIF 
ENDIF 

OutDirPeriod = pOut + '\' + m.gcperiod

IF !fso.FolderExists(OutDirPeriod)
 NewDirr = fso.CreateFolder(OutDirPeriod)
 RELEASE NewDirr
ENDIF 

FinFile = 'f13'+m.qcod

fso.CopyFile(pTempl+'\f13qq.dbf', OutDirPeriod+'\'+finfile+'.dbf')

=OpenFile(OutDirPeriod+'\'+finfile, 'finfile', 'shared')

SELECT aisoms
m.n_rec = 0
SCAN FOR !DELETED()
SCATTER MEMVAR 
 m.s_pred  = m.s_pred + m.s_lek
 m.sum_flk = m.sum_flk + m.ls_flk
 
 m.IsPilot  = IIF(!EMPTY(pr4.mcod) OR INLIST(m.lpuid,4708,4963), .T., .F.)
 m.IsPilotS = IIF(!EMPTY(pr4st.mcod), .T., .F.)

 m.IsTpn = .f.
 IF USED('lputpn')
  m.IsTpn = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)
 ENDIF 

 m.n_rec      = m.n_rec + 1

 m.recid      = PADL(m.n_rec,6,'0')
 m.lpu_id     = m.lpuid
 m.info_type  = 'M'
 m.period     = m.gcperiod
 m.s_pred_pf  = (pr4.s_own + pr4.s_bad)+(pr4st.s_own + pr4st.s_bad) && ok!
 m.s_calc_pf  = pr4.finval + pr4st.finval && ok!
 *m.s_pr_opl   = (m.s_pred+m.s_lek) - m.sum_flk && ok!
 m.s_pr_opl   = m.s_pred - m.sum_flk && ok!
 m.s_pr_p     = pr4.s_own + pr4st.s_own && ok!
 m.s_pr_pno   = pr4.s_others + pr4st.s_others && ok! pr4st.s_others таких больше нет!
 m.s_pr_outp  = pr4.s_guests + pr4st.s_guests && ok! pr4st.s_guests таких больше нет!

 m.s_pr_nopf  = mag.col22+mag.col23 && s_dop+sum_st+sum_dst+sum_vmp && переменная поменяла свой смысл!!! 

 m.s_pr_nop   = pr4.s_empty + pr4st.s_empty
 
 m.s_fact    = 0
 m.s_profusl = 0
 m.s_dayhos  = 0
 m.s_rest_pf = 0
  
 m.period_e  = ''
 m.et        = ''
 m.s_expert1 = 0
 m.kor_exp   = 0
 
 m.s_fine     = 0
  
 m.s_dop1 = m.s_avans2

 *IF m.qcod<>'I3'
 * m.t_t        = m.s_avans
 * m.s_avans    = m.s_pr_avans
 * m.s_pr_avans = m.t_t
 *ENDIF 

 IF EMPTY(pr4.mcod) AND EMPTY(pr4st.mcod)
  *m.s_pr_avans = 0
  m.s_k_perech = IIF(m.s_pr_opl - m.s_avans - m.dolg_b>0, ;
   m.s_pr_opl - m.s_avans - m.dolg_b, 0)
  m.itog      = 0
  m.s_pred_pf = 0
  m.s_pr_nopf = 0
  m.kor_exp   = 0

  IF (s_pred>0 OR s_avans>0) OR m.mcod='0371001'
   INSERT INTO finfile FROM MEMVAR 
  ENDIF 
 ELSE 
  m.s_k_perech = m.s_calc_pf - m.s_avans - m.s_pr_pno + m.s_pr_outp + ;
  m.s_pr_nop + m.s_pr_nopf
  m.itog = m.s_calc_pf - IIF(m.IsTpn, m.s_pr_avans, m.s_avans) - m.s_pr_pno + m.s_pr_outp + ;
  m.s_pr_nop

  INSERT INTO finfile FROM MEMVAR 

 ENDIF 

 IF m.llIsMeFiles
  IF SEEK(m.lpu_id, 'svmee')
   SELECT * FROM svmee WHERE lpuid=m.lpu_id ORDER BY e_period,et INTO CURSOR memcod
   kol_exp = RECCOUNT('memcod')
   FOR n_exp=1 TO kol_exp

    m.n_rec      = m.n_rec + 1
    m.recid      = PADL(m.n_rec,6,'0')
    m.period_e   = memcod.e_period
    m.et         = et
    m.et = IIF(SEEK(m.et, 'mee2mgf'), mee2mgf.mgf_et, m.et)
    m.s_expert1  = memcod.sexp
    m.s_fine     = memcod.s_sank
    IF m.IsPilot OR m.IsPilotS
     IF m.IsTpn = .f.
      m.kor_exp    = m.s_expert1
     ELSE 
      m.kor_exp    = IIF(VAL(m.period_e)>201406, memcod.stpn, 0)
     ENDIF 
    ELSE 
     m.kor_exp    = 0
    ENDIF 
    m.s_k_perech = m.s_k_perech - m.s_expert1
    m.itog       = m.itog  - IIF(m.IsTpn,  m.kor_exp, m.s_expert1)
    m.itog       = IIF(m.itog>=0, m.itog, 0)

    INSERT INTO finfile ;
     (recid, lpu_id, info_type, period, period_e, et, s_expert1, kor_exp, s_fine) ;
    VALUES ;
     (m.recid, m.lpu_id, 'M', m.gcperiod, m.period_e, m.et, -1*m.s_expert1, IIF(m.IsPilot, -1*m.kor_exp, 0), m.s_fine) 
     
    UPDATE finfile SET s_k_perech=m.s_k_perech, itog=m.itog ;
     WHERE STR(lpu_id,4)=STR(m.lpu_id,4) AND s_pred>0
    IF _tally>0
    ENDIF 

    SKIP IN memcod
   ENDFOR 
   USE IN memcod
  ENDIF 
 ENDIF 

ENDSCAN 

SELECT SUM(s_k_perech) as s_k_perech, SUM(itog) as itog, SUM(s_expert1) as s_expert1 ;
 FROM finfile INTO CURSOR bnm
m.s_k_perech = bnm.s_k_perech
m.itog = bnm.itog
m.s_expert1 = bnm.s_expert1
USE IN bnm

SELECT aisoms
SET RELATION OFF INTO SprLpu
SET RELATION OFF INTO pr4
SET RELATION OFF INTO pr4st
SET RELATION OFF INTO mag

USE IN mag
USE IN SprLpu
USE IN pr4
USE IN pr4st
USE IN aisoms 
IF USED('lputpn')
 USE IN lputpn
ENDIF 

USE IN finfile
IF USED('svmee')
 USE IN svmee
ENDIF 
IF USED('pplais')
 USE IN pplais
ENDIF 
USE IN mee2mgf


MESSAGEBOX(CHR(13)+CHR(10)+'SUM(s_k_perech): '+TRANSFORM(m.s_k_perech, '9999999999.99')+CHR(13)+CHR(10)+;
 'SUM(itog)      : '+TRANSFORM(m.itog, '9999999999.99')+CHR(13)+CHR(10)+;
 'SUM(s_expert1) : '+TRANSFORM(m.s_expert1, '9999999999.99')+CHR(13)+CHR(10),0+64,'')

RETURN 