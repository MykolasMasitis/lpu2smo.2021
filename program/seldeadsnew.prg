PROCEDURE seldeadsnew
 IF MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ������?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 CREATE CURSOR curdeads (nrec i AUTOINC, period c(7), lpuid i(4), mcod c(7), sn_pol c(25), c_i c(30), ;
  fam c(25), im c(25), ot c(25), dr d, w n(1), d_u d, ds c(6), ds_n c(150), ds_2 c(6), ds_2_n c(150), ds_3 c(6), ds_3_n c(150),;
  ds_pat c(6), ds_pat_n c(150), ds_2_pat c(6), ds_2_pat_n c(150), ds_3_pat c(6), ds_3_pat_n c(150), otd c(4), otdname c(100), ;
  pcod c(10), docname c(100), cod n(6), codname c(100), k_u n(3), n_kd n(3), tip c(1),;
  s_all n(11,2), d_beg d, d_end d, ishod n(3),lpuname c(40), ischked l,;
  osn230 c(6), osn230_n c(150), koeff n(4,2), straf n(4,2), s_1 n(11,2), s_2 n(11,2), codho c(14), honame c(100), usr c(6), usrname c(150),;
  d_exp d, docexp c(6),dschet d, nschet c(7), n_dog c(16))
 INDEX on nrec TAG nrec 
 CREATE CURSOR curpats (period c(7), lpuid i(4), mcod c(7), sn_pol c(25), c_i c(30), ;
  fam c(25), im c(25), ot c(25), dr d, w n(1), d_u d, ds c(6), ds_n c(150),ds_2 c(6), ds_2_n c(150), ds_3 c(6), ds_3_n c(150),;
  otd c(4), otdname c(100), pcod c(10), docname c(100),;
  cod n(6), codname c(100), k_u n(3), s_all n(11,2), d_beg d, d_end d, lpuname c(40), ischked l)
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 FOR lnmonth=1 TO m.tmonth
  m.lcperiod = STR(tYear,4)+PADL(lnmonth,2,'0')
  m.lpath = pbase+'\'+m.lcperiod
  IF !fso.FolderExists(m.lpath)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.lpath+'\aisoms.dbf')
   LOOP 
  ENDIF 
  
  WAIT m.lcperiod+'...' WINDOW NOWAIT 
  IF OpenFile(m.lpath+'\nsi'+'\tarifn', 'tarif', 'shar', 'cod')>0
   IF USED('tarif')
    USE IN tarif 
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.lpath+'\nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
   USE IN tarif 
   IF USED('sprlpu')
    USE IN sprlpu
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.lpath+'\nsi'+'\mkb10', 'mkb', 'shar', 'ds')>0
   USE IN tarif 
   USE IN sprlpu
   IF USED('mkb')
    USE IN mkb
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.lpath+'\nsi'+'\sookodxx', 'sookod', 'shar', 'osn230')>0
   USE IN tarif 
   USE IN sprlpu
   USE IN mkb
   IF USED('sookod')
    USE IN sookod
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.pcommon+'\lpudogs', 'lpudogs', 'shar', 'lpu_id')>0
   USE IN tarif 
   USE IN sprlpu
   USE IN mkb
   USE IN sookod
   IF USED('lpudogs')
    USE IN lpudogs
   ENDIF 
   LOOP 
  ENDIF 
*  IF OpenFile(pcommon+'\users', 'ussr', 'shar', 'name')>0
*   USE IN sookod
*   USE IN tarif 
*   USE IN sprlpu
*   USE IN mkb
*   IF USED('users')
*    USE IN users
*   ENDIF 
*   LOOP 
*  ENDIF 

  =seldeadsone(m.lpath)

  USE IN tarif 
  USE IN sprlpu
  USE IN mkb
  USE IN sookod
  USE IN lpudogs
*  USE IN ussr

  WAIT CLEAR 

 NEXT 

 =OpenFile(pcommon+'\hopff_xx', 'hopff', 'shar', 'codho')
 SELECT curdeads
 SET RELATION TO sn_pol INTO curpats
 SET RELATION TO codho INTO hopff ADDITIVE 
 REPLACE ALL ds_pat WITH curpats.ds, ds_2_pat WITH curpats.ds_2, ds_3_pat WITH curpats.ds_3,;
  ds_pat_n WITH curpats.ds_n, ds_2_pat_n WITH curpats.ds_2_n, ds_3_pat_n WITH curpats.ds_3_n,;
  honame WITH ALLTRIM(hopff.name)
 SET RELATION OFF INTO curpats
 SET RELATION OFF INTO hopff 
 IF USED('hopff')
  USE IN hopff
 ENDIF 
 m.dotname = ptempl+'\sh_new.xls'
 m.docname = pmee+'\DList'+m.qcod+PADL(DAY(DATE()),2,'0')+PADL(MONTH(DATE()),2,'0')
 IF fso.FileExists(m.docname+'.xls')
  fso.DeleteFile(m.docname+'.xls')
 ENDIF 
 m.llResult = X_Report(m.dotname, m.docname+'.xls', .T.)
 USE IN curdeads
	
RETURN 

FUNCTION seldeadsone(m.lpath)
 PRIVATE m.llcpath
 m.llcpath = m.lpath
 IF OpenFile(m.llcpath+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 SELECT aisoms
 SCAN 
  m.lpuid = lpuid
  m.mcod = mcod
  IF !fso.FolderExists(m.llcpath+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpFiles()<0
   =ClFiles()
   LOOP 
  ENDIF 

  m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.fullname, '')
  m.n_dog = IIF(SEEK(m.lpuid,'lpudogs'), lpudogs.dogs,'')
 
  SELECT talon 
  SET RELATION TO sn_pol INTO people
  SET RELATION TO pcod INTO doctor ADDITIVE 
  SET RELATION TO otd INTO otdel ADDITIVE 
  SET RELATION TO recid INTO error ADDITIVE 
  SET RELATION TO recid INTO merror ADDITIVE 
  SET RELATION TO cod INTO tarif ADDITIVE 
  IF USED('ho')
   SET RELATION TO sn_pol+c_i+PADL(cod,6,"0") INTO ho ADDITIVE 
  ENDIF 
  SCAN 
   IF !EMPTY(error.rid)
    LOOP 
   ENDIF 
   IF !(d_type='5' OR tip='5' OR IsPat(cod))
    LOOP 
   ENDIF 
   
   m.sn_pol = sn_pol
   m.c_i    = c_i
   m.fam    = people.fam
   m.im     = people.im
   m.ot     = people.ot
   m.dr     = people.dr
   m.w      = people.w
   m.d_beg  = people.d_beg
   m.d_end  = people.d_end
   m.d_u    = d_u
   m.ds     = ds
   m.ds_n   = IIF(SEEK(m.ds, 'mkb'), ALLTRIM(mkb.name_ds), '')
   m.ds_2   = ds_2
   m.ds_2_n = IIF(SEEK(m.ds_2, 'mkb'), ALLTRIM(mkb.name_ds), '')
   m.ds_3   = ds_3
   m.ds_3_n = IIF(SEEK(m.ds_3, 'mkb'), ALLTRIM(mkb.name_ds), '')
   m.otd    = otd
   m.otdname = ALLTRIM(otdel.name)
   m.pcod   = pcod
   m.docname = ALLTRIM(doctor.fam)+' '+ALLTRIM(doctor.im)+' ' +ALLTRIM(doctor.ot)
   m.cod    = cod
   m.codname = tarif.comment
   m.k_u    = k_u
   m.s_all  = s_all
   m.ishod  = ishod
   m.ischked = IIF(EMPTY(merror.recid), .f., .t.)
   m.osn230 = merror.osn230
   m.koeff  = merror.koeff
   m.straf  = merror.straf
   m.s_1    = merror.s_1
   m.s_2    = merror.s_2
   m.usr    = merror.usr
   m.usrname = IIF(SEEK(m.usr, 'users', 'name'), ALLTRIM(users.fam)+' '+ALLTRIM(users.im)+' '+ALLTRIM(users.ot), '')
   m.d_exp  = merror.d_exp
   m.docexp = merror.docexp
   m.codho = IIF(USED('ho'), IIF(!EMPTY(ho.codho), ho.codho, ''), '')
   m.nschet = m.lcperiod
   m.dschet = CTOD('05.'+PADL(MONTH(m.d_u),2,'0')+'.'+STR(YEAR(m.d_u),4))
   m.n_kd   = n_kd
   m.tip    = tip
   m.osn230_n = IIF(SEEK(m.osn230, 'sookod'), sookod.f_naim, '')
   
   IF !IsPat(m.cod)
    INSERT INTO curdeads (period,lpuid,mcod,sn_pol,c_i,fam,im,ot,dr,w,;
     d_u,ds,ds_n,ds_2,ds_2_n,ds_3,ds_3_n,otd,otdname,pcod,docname,cod,codname,k_u, s_all, ishod,d_beg, d_end,lpuname,ischked,;
     osn230,osn230_n,koeff,straf,s_1,s_2,codho,usr,usrname,d_exp,docexp,dschet,nschet,n_kd,tip,n_dog) VALUES ;
     (m.lcperiod,m.lpuid,m.mcod,m.sn_pol,m.c_i,m.fam,m.im,m.ot,m.dr,m.w,;
      m.d_u,m.ds,m.ds_n,m.ds_2,m.ds_2_n,m.ds_3,m.ds_3_n,m.otd,m.otdname,m.pcod,m.docname,m.cod,m.codname,m.k_u, m.s_all,m.ishod, ;
      m.d_beg, m.d_end,m.lpuname,m.ischked,m.osn230,m.osn230_n,m.koeff,m.straf,m.s_1,m.s_2,m.codho,m.usr,m.usrname,m.d_exp,m.docexp,;
      m.dschet,m.nschet,m.n_kd,m.tip,m.n_dog) 
   ELSE 
    INSERT INTO curpats (period,lpuid,mcod,sn_pol,c_i,fam,im,ot,dr,w,;
     d_u,ds,ds_n,ds_2,ds_2_n,ds_3,ds_3_n,otd,otdname,pcod,docname,cod,codname,k_u, s_all, d_beg, d_end,lpuname,ischked) VALUES ;
     (m.lcperiod,m.lpuid,m.mcod,m.sn_pol,m.c_i,m.fam,m.im,m.ot,m.dr,m.w,;
      m.d_u,m.ds,m.ds_n,m.ds_2,m.ds_2_n,m.ds_3,m.ds_3_n,m.otd,m.otdname,m.pcod,m.docname,m.cod,m.codname,m.k_u, m.s_all, ;
      m.d_beg, m.d_end,m.lpuname,m.ischked) 
   ENDIF 

  ENDSCAN 
  SET RELATION OFF INTO doctor
  SET RELATION OFF INTO otdel
  SET RELATION OFF INTO merror
  SET RELATION OFF INTO error
  SET RELATION OFF INTO people
  SET RELATION OFF INTO tarif
  IF USED('ho')
   SET RELATION OFF INTO ho
  ENDIF 

  =ClFiles()
 
  SELECT aisoms

 ENDSCAN 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 

RETURN 

FUNCTION OpFiles()
 IF OpenFile(m.llcpath+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
  RETURN -1 
 ENDIF 
 IF OpenFile(m.llcpath+'\'+m.mcod+'\talon', 'talon', 'shar')>0
  RETURN -1 
 ENDIF 
 IF OpenFile(m.llcpath+'\'+m.mcod+'\otdel', 'otdel', 'shar', 'iotd')>0
  RETURN -1 
 ENDIF 
 IF OpenFile(m.llcpath+'\'+m.mcod+'\doctor', 'doctor', 'shar', 'pcod')>0
  RETURN -1 
 ENDIF 
 IF OpenFile(m.llcpath+'\'+m.mcod+'\e'+m.mcod, 'error', 'shar', 'rid')>0
  RETURN -1 
 ENDIF 
 IF OpenFile(m.llcpath+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar', 'recid')>0
  RETURN -1 
 ENDIF 
 IF fso.FileExists(m.llcpath+'\'+m.mcod+'\ho'+m.qcod+'.dbf')
  IF OpenFile(m.llcpath+'\'+m.mcod+'\ho'+m.qcod, 'ho', 'shar', 'unik')>0
   RETURN -1 
  ENDIF 
 ENDIF 
RETURN 0

FUNCTION ClFiles()
 IF USED('people')
  USE IN people
 ENDIF
 IF USED('talon')
  USE IN talon
 ENDIF 
 IF USED('otdel')
  USE IN otdel
 ENDIF 
 IF USED('doctor')
  USE IN doctor
 ENDIF 
 IF USED('error')
  USE IN error
 ENDIF 
 IF USED('merror')
  USE IN merror
 ENDIF 
 IF USED('ho')
  USE IN ho
 ENDIF 
RETURN 