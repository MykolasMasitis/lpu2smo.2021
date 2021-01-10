PROCEDURE Sel50PV2

 IF MESSAGEBOX("ÎÒÎÁÐÀÒÜ ÊÎÐÎÒÊÈÅ È ÄËÈÍÍÛÅ"+CHR(13)+CHR(10)+"ÃÎÑÏÈÒÀËÈÇÀÖÈÈ?"+CHR(13)+CHR(10),;
  4+32,'ÍÎÂÛÉ ÂÀÐÈÀÍÒ')=7
  RETURN 
 ENDIF 

 m.dotname = pTempl+'\sh50ver2.xls'
 m.doñname = pTempl+'\sh50ver2.xls'
 IF !fso.FileExists(m.dotname)
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË ØÀÁËÎÍÀ '+CHR(13)+CHR(10)+;
   m.dotname,0+64,'')
  RETURN 
 ENDIF 

 m.RepDir = fso.GetParentFolderName(pbin) + '\REPS'
 IF !fso.FolderExists(m.RepDir)
  fso.CreateFolder(m.RepDir)
 ENDIF 

 m.BookName = m.repdir+'\sh50ver2'+m.gcperiod
 m.IsOpDoc = IsOpenExcelDoc('sh50ver2'+m.gcperiod)
 IF m.IsOpDoc
  IF !CloseExcelDoc('sh50ver2'+m.gcperiod)
   MESSAGEBOX('ÔÀÉË '+'sh50ver2'+m.gcperiod+' ÎÒÊÐÛÒ!',0+64,'')
   RETURN .f. 
  ENDIF 
 ENDIF 
  
 CREATE CURSOR allguests (nrec i AUTOINC, period c(7), lpuid i(4), mcod c(7), sn_pol c(25), c_i c(30), ;
   fam c(25), im c(25), ot c(25), dr d, w n(1),d_u d, ds c(6), otd c(4), pcod c(10), cod n(6), tip c(1), k_u n(3), n_kd n(3),;
   s_all n(11,2), d_beg d, d_end d, codname c(40), lpuname c(40), q c(2),;
   isexp l, et c(1), osn230 c(5), s_1 n(11,2), s_2 n(11,2))

 CREATE CURSOR curbooks (bkname c(13))

 FOR lnmonth=1 TO 12
  m.lcperiod = STR(tYear,4)+PADL(lnmonth,2,'0')
  m.lpath = pbase+'\'+m.lcperiod
  IF !fso.FolderExists(m.lpath)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.lpath+'\aisoms.dbf')
   LOOP 
  ENDIF 
  
  WAIT m.lcperiod+'...' WINDOW NOWAIT 
  =SelAlienOne(m.lpath)
  WAIT CLEAR 

 ENDFOR 
 
 m.docname = m.repdir+'\sh50ver2'+m.gcperiod+'.XLS'
 SELECT allguests 
  IF RECCOUNT('allguests')>0
  m.llResult = X_Report(m.dotname, m.docname, .t.)
  ENDIF 
 USE IN allguests

 IF 3=2
 PUBLIC oExcel AS Excel.Application
 WAIT "Çàïóñê MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 
 
 oexcel.SheetsInNewWorkbook = 1
 AllBook = oExcel.Workbooks.Add()
 SELECT curbooks
 SCAN 
  m.lpucod = SUBSTR(bkname,3,7)
  m.bkname = m.RepDir+'\'+bkname
  IF !fso.FileExists(m.bkname)
   LOOP 
  ENDIF 
  OneBook  = oExcel.Workbooks.Add(m.bkname)
  OneSheet = OneBook.WorkSheets(1).Copy(,AllBook.ActiveSheet)
  OneBook.Close
  fso.DeleteFile(m.bkname)
  AllBook.ActiveSheet.Name = m.lpucod
 ENDSCAN 
 USE IN curbooks 

 IF fso.FileExists(m.repdir+'\sh'+m.gcperiod+'.xls')
  fso.DeleteFile(m.repdir+'\sh'+m.gcperiod+'.xls')
 ENDIF  
 AllBook.SaveAs(m.repdir+'\sh'+m.gcperiod,18)
 oExcel.Visible=.t.
 ENDIF 

RETURN 


FUNCTION SelAlienOne(m.lpath)

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
  m.mcod  = mcod
  IF INT(VAL(SUBSTR(m.mcod,3,2)))<41
   LOOP 
  ENDIF 
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
  IF !fso.FileExists(m.llcpath+'\nsi'+'\tarifn'+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\nsi'+'\sprlpuxx'+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(m.llcpath+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llcpath+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llcpath+'\'+m.mcod+'\e'+m.mcod, 'error', 'shar', 'rid')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('error')
    USE IN error
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llcpath+'\nsi'+'\tarifn', 'tarif', 'shar', 'cod')>0
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('error')
    USE IN error
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llcpath+'\nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
   IF USED('sprlpu')
    USE IN sprlpu
   ENDIF 
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('error')
    USE IN error
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llcpath+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar', 'recid')>0
   IF USED('sprlpu')
    USE IN sprlpu
   ENDIF 
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('error')
    USE IN error
   ENDIF 
   IF USED('merror')
    USE IN merror
   ENDIF 
   LOOP 
  ENDIF 
  
*  WAIT m.mcod WINDOW NOWAIT 
  
*  m.docname = m.repdir+'\sh'+m.mcod+'.XLS'
*  INSERT INTO curbooks (bkname) VALUES ('sh'+m.mcod+'.XLS')
*  IF fso.FileExists(m.docname)
*   fso.DeleteFile(m.docname)
*  ENDIF 

  m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.name, '')
 
  SELECT talon 
  SET RELATION TO sn_pol INTO people
  SET RELATION TO recid INTO error ADDITIVE 
  SET RELATION TO recid INTO merror ADDITIVE 
  SET RELATION TO cod INTO tarif ADDITIVE 
  SCAN 
   IF !EMPTY(error.rid)
    LOOP 
   ENDIF 
   m.cod = cod
   IF !IsMes(m.cod)
    LOOP 
   ENDIF 
   m.k_u   = k_u
   m.n_kd  = tarif.n_kd
   m.koeff = ROUND(m.k_u/m.n_kd,1)
   
   IF BETWEEN(m.koeff,0.5,1.5)
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
   m.otd    = otd
   m.pcod   = pcod
   m.cod    = cod
   m.tip    = tip
   m.k_u    = k_u
*   m.n_kd   = n_kd
   m.s_all  = s_all
   m.isexp  = IIF(!EMPTY(merror.osn230), .t., .f.)
   m.et     = merror.et
   m.osn230 = merror.osn230
   m.s_1    = merror.s_1
   m.s_2    = merror.s_2
   
   m.codname = IIF(SEEK(m.cod, 'tarif'), tarif.comment, '')

   INSERT INTO allguests (period,lpuid,mcod,sn_pol,c_i,fam,im,ot,dr,w,;
    d_u,ds,otd,pcod, cod, tip, k_u, n_kd, s_all, d_beg, d_end, codname,lpuname,;
    isexp,et,osn230,s_1,s_2) VALUES ;
    (m.gcperiod,m.lpuid,m.mcod,m.sn_pol,m.c_i,m.fam,m.im,m.ot,m.dr,m.w,;
    m.d_u,m.ds,m.otd,m.pcod,m.cod,m.tip, m.k_u, m.n_kd, m.s_all, ;
    m.d_beg, m.d_end, m.codname,m.lpuname,m.isexp,m.et,m.osn230,m.s_1,m.s_2) 

  ENDSCAN 

  SET RELATION OFF INTO error
  SET RELATION OFF INTO merror
  SET RELATION OFF INTO people
  SET RELATION OFF INTO tarif 
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('error')
   USE IN error
  ENDIF 
  IF USED('merror')
   USE IN merror
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  
 
  SELECT aisoms
  
*  WAIT CLEAR 

 ENDSCAN 
 USE IN aisoms

RETURN 

