FUNCTION rep_mp(para1)
 m.mcod = para1
 IF MESSAGEBOX('ВЫ ХОТИТЕ СФОРМИРОВАТЬ'+CHR(13)+CHR(10)+'ОТЧЕТ ПО МП?',4+32,'МО '+m.mcod)=7
  RETURN
 ENDIF 
 IF !fso.FileExists(pTempl+'\rep_mo.xls')
  MESSAGEBOX('ОТСУТСТВУЕТ ШАБЛОН ОТЧЕТА'+CHR(13)+CHR(10)+UPPER(pTempl+'\rep_mo.xls'),0+64,'')
  RETURN 
 ENDIF 
 
 IF fso.FileExists(pbase+'\rep_mo'+STR(tYear,4)+'.dbf')
  fso.DeleteFile(pbase+'\rep_mo'+STR(tYear,4)+'.dbf')
 ENDIF 

 CREATE CURSOR curdata (nrec i AUTOINC, paz n(7), obr n(7), pos n(7), pos_s n(11,2),;
  parakl n(7), n37047 n(7), s37047 n(11,2), parakl_s n(11,2), usls n(7), npaz n(7), nsch n(7), stpaz n(7), k_dn n(7), stpaz_s n(11,2),;
  vmp n(7), vmpk_dn n(7), vmp_s n(11,2), ecos n(7), ecos_s n(11,2), dstpaz n(7), k_dnst n(7), dstpaz_s n(11,2),;
  paz_zab n(7), sum_zab n(11,2), paz_02 n(7), sum_02 n(11,2), paz_prof n(7), sum_prof n(11,2))
 INDEX on nrec TAG nrec
 SET ORDER TO nrec 
  
 FOR m.nmonth=1 TO m.tmonth
  m.lcperiod = LEFT(m.gcperiod,4)+PADL(m.nmonth,2,'0')
  m.lcmonth = PADL(m.nmonth,2,'0')

  =rep_one(m.lcperiod)
  
 ENDFOR 

 m.llResult = X_Report(pTempl+'\rep_mo.xls', pBase+'\'+m.gcPeriod+'\'+m.mcod+'\rep_mo.xls', .T.)

 USE IN curdata
 
RETURN 

FUNCTION rep_one(para01)
 m.lcperiod = para01
 m.lcmonth  = SUBSTR(para01,5,2)
 
 IF !fso.FolderExists(pbase+'\'+m.lcperiod +'\'+m.mcod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.lcperiod +'\'+m.mcod+'\talon.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.lcperiod +'\'+m.mcod+'\talon', 'talon', 'shar')>0
  IF USED('talon')
   USE IN talon 
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.lcperiod +'\'+m.mcod+'\e'+m.mcod, 'err', 'shar', 'rid')>0
  IF USED('err')
   USE IN err
  ENDIF 
  USE IN talon 
  RETURN 
 ENDIF 

 CREATE CURSOR curobr (period c(7), sn_pol c(25))
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol
  
 CREATE CURSOR curpazst (period c(7), c_i c(30))
 INDEX ON c_i TAG c_i
 SET ORDER TO c_i

 CREATE CURSOR curpazvmp (period c(7), c_i c(30))
 INDEX ON c_i TAG c_i
 SET ORDER TO c_i

 CREATE CURSOR curpazdst (period c(7), c_i c(30))
 INDEX ON c_i TAG c_i
 SET ORDER TO c_i

 CREATE CURSOR curpazeco (period c(7), c_i c(30))
 INDEX ON c_i TAG c_i
 SET ORDER TO c_i

 CREATE CURSOR curpaz02 (period c(7), sn_pol c(25))
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR curpazzab (period c(7), sn_pol c(25))
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR curpazprof (period c(7), sn_pol c(25))
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol

  m.obr      = 0
  m.obr_s    = 0
  m.pos      = 0
  m.pos_s    = 0
  m.parakl   = 0
  m.parakl_s = 0
  m.usls     = 0
  
  m.stpaz   = 0
  m.k_dn    = 0
  m.stpaz_s = 0
  m.vmp     = 0
  m.vmp_s   = 0
  m.vmpk_dn = 0
  m.dstpaz  = 0
  m.k_dnst  = 0
  m.dstpaz_s=0
  
  m.ecos   = 0
  m.ecos_s = 0
  m.n37047 = 0
  m.s37047 = 0
  
  m.sum_02   = 0
  m.sum_prof = 0
  m.sum_zab  = 0
  
  SELECT talon 
  SET RELATION TO recid INTO err
  SCAN
   IF !EMPTY(err.c_err)
    LOOP 
   ENDIF 
   
   m.c_i    = c_i
   m.sn_pol = sn_pol
   m.cod    = cod 
   m.k_u    = k_u
   m.s_all  = s_all + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
   m.p_cel  = p_cel
   
   IF IsUsl(m.cod)
    IF !SEEK(m.sn_pol, 'curobr', 'sn_pol')
     INSERT INTO curobr (period, sn_pol) VALUES (m.lcperiod ,m.sn_pol) && кол-во пациентов!
    ENDIF 
    
    DO CASE 
     CASE m.p_cel = '1.1'
      IF !SEEK(m.sn_pol, 'curpaz02', 'sn_pol')
       INSERT INTO curpaz02 (period, sn_pol) VALUES (m.lcperiod ,m.sn_pol) && кол-во пациентов!
      ENDIF 
      m.sum_02 = m.sum_02 + m.s_all
     CASE INLIST(m.p_cel, '1.3','2.1','2.2','2.3','2.5','2.6','3.1')
      IF !SEEK(m.sn_pol, 'curpazprof', 'sn_pol')
       INSERT INTO curpazprof (period, sn_pol) VALUES (m.lcperiod ,m.sn_pol) && кол-во пациентов!
      ENDIF 
      m.sum_prof = m.sum_prof + m.s_all
     OTHERWISE 
      IF !SEEK(m.sn_pol, 'curpazzab', 'sn_pol')
       INSERT INTO curpazzab (period, sn_pol) VALUES (m.lcperiod ,m.sn_pol) && кол-во пациентов!
      ENDIF 
      m.sum_zab = m.sum_zab + m.s_all
    ENDCASE 

    IF SUBSTR(PADL(m.cod,6,'0'),3,1) = '1' && посещение
     m.pos = m.pos + m.k_u
     m.pos_s = m.pos_s + m.s_all
     IF SUBSTR(PADL(m.cod,6,'0'),3,1)='1' AND SUBSTR(PADL(m.cod,6,'0'),6,1)='1'&& обращение
      m.obr   = m.obr + m.k_u
      m.obr_s = m.obr_s + m.s_all
     ENDIF 
    ELSE && параклиника
     m.parakl   = m.parakl + m.k_u
     m.parakl_s = m.parakl_s + m.s_all
    ENDIF 
    
    IF m.cod = 37047
     m.n37047 = m.n37047 + m.k_u
     m.s37047 = m.s37047 + m.s_all
    ENDIF 

    m.usls = m.usls + m.k_u
   ENDIF 
   
   IF IsDst(m.cod)
    IF !SEEK(m.c_i, 'curpazdst', 'c_i')
     INSERT INTO curpazdst (period, c_i) VALUES (m.lcperiod, m.c_i) && кол-во пациентов!
    ENDIF 
    m.k_dnst   = m.k_dnst + m.k_u
    m.dstpaz_s = m.dstpaz_s + m.s_all

    IF IsEKO(m.cod)
     IF !SEEK(m.c_i, 'curpazeco', 'c_i')
      INSERT INTO curpazeco (period, c_i) VALUES (m.lcperiod , m.c_i) && кол-во пациентов!
     ENDIF 
     m.ecos_s = m.ecos_s + m.s_all
    ENDIF 

   ENDIF 
   
   IF IsGsp(m.cod)
    IF !SEEK(m.c_i, 'curpazst', 'c_i')
     INSERT INTO curpazst (period, c_i) VALUES (m.lcperiod, m.c_i) && кол-во пациентов!
    ENDIF 
    IF m.IsVMP(m.cod) AND !SEEK(m.c_i, 'curpazvmp', 'c_i')
     INSERT INTO curpazvmp (period, c_i) VALUES (m.lcperiod ,m.c_i) && кол-во пациентов!
    ENDIF 
    m.k_dn    = m.k_dn + m.k_u
    m.stpaz_s = m.stpaz_s + m.s_all
    m.vmp_s   = m.vmp_s + IIF(m.IsVmp(m.cod), m.s_all, 0)
    m.vmpk_dn    = m.vmpk_dn + IIF(m.IsVMP(m.cod), m.k_u, 0)
   ENDIF 
   

  ENDSCAN 
  m.paz      = RECCOUNT('curobr')
  m.stpaz    = RECCOUNT('curpazst')
  m.dstpaz   = RECCOUNT('curpazdst')
  m.ecos     = RECCOUNT('curpazeco')
  m.vmp      = RECCOUNT('curpazvmp')
  m.paz_02   = RECCOUNT('curpaz02')
  m.paz_zab  = RECCOUNT('curpazzab')
  m.paz_prof = RECCOUNT('curpazprof')

  SET RELATION OFF INTO err
  USE IN talon 
  USE IN err
  USE IN curobr
  USE IN curpazst
  USE IN curpazdst
  USE IN curpazeco
  USE IN curpazvmp
  USE IN curpaz02
  USE IN curpazzab
  USE IN curpazprof
  
  
  IF RECCOUNT('curdata')<=0
   INSERT INTO curdata FROM MEMVAR 
  ELSE 
   UPDATE curdata SET paz=paz+m.paz, obr=obr+m.obr, pos=pos+m.pos, pos_s=pos_s+m.pos_s, parakl=parakl+m.parakl,;
     n37047=n37047+m.n37047, s37047=s37047+m.s37047,;
     parakl_s=parakl_s+m.parakl_s, usls=usls+m.usls, ;
     stpaz=stpaz+m.stpaz, k_dn=k_dn+m.k_dn, stpaz_s=stpaz_s+m.stpaz_s, vmp=vmp+m.vmp, vmpk_dn=vmpk_dn+m.vmpk_dn,;
     vmp_s=vmp_s+m.vmp_s, ecos=ecos+m.ecos, ecos_s=ecos_s+m.ecos_s, ;
     sum_02 = sum_02 + m.sum_02, sum_prof = sum_prof + m.sum_prof, sum_zab = sum_zab + m.sum_zab,;
     paz_02 = paz_02 + m.paz_02, paz_prof = paz_prof + m.paz_prof, paz_zab = paz_zab + m.paz_zab
     
     
  ENDIF 
   

RETURN 