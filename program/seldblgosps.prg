PROCEDURE seldblgosps
 IF MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ��������������?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 CREATE CURSOR curgosps (period c(7), lpuid i(4), mcod c(7), sn_pol c(25), c_i c(30), ;
  fam c(25), im c(25), ot c(25), dr d, w n(1),d_u d, ds c(6), dss c(3), otd c(4), pcod c(10), ;
  cod n(6), k_u n(5),tip c(1), n_kd n(3),s_all n(11,2), d_beg d, d_end d, codname c(40), lpuname c(40), ischked l)
 INDEX on c_i TAG c_i
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO c_i 

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
  =seldblgosp(m.lpath)
  WAIT CLEAR 

 ENDFOR 

 CREATE CURSOR curmcod (mcod c(7), lpuname c(40))
 INDEX on mcod TAG mcod
 SET ORDER TO mcod 

 outfile = pmee+'\sldblgosps'
 CREATE TABLE &outfile (period c(7), lpuid i(4), mcod c(7), sn_pol c(25), c_i c(30), ;
  fam c(25), im c(25), ot c(25), dr d, w n(1), d_u d, ds c(6), dss c(3), cod n(6), k_u n(5),;
  otd c(4), pcod c(10), d_u1 d, d_u2 d, tip c(1), n_kd n(3),s_all n(11,2), d_beg d, d_end d, ;
  codname c(40), lpuname c(40), ischked l)
 USE 

 =OpenFile(outfile, 'outfl', 'shar')
 
 SELECT sn_pol, MIN(d_u) as d_u1, MAX(d_u) as d_u2 FROM curgosps ;
  GROUP BY sn_pol, dss HAVING coun(*)>1 INTO CURSOR cur_tmpl
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 
 SELECT curgosps
 SET ORDER TO sn_pol
 SET RELATION TO sn_pol INTO cur_tmpl
 SCAN 
  IF EMPTY(cur_tmpl.sn_pol)
   LOOP 
  ENDIF 

  SCATTER MEMVAR 

  m.d_u1 = cur_tmpl.d_u1
  m.d_u2 = cur_tmpl.d_u2
  
  IF m.d_u2-m.d_u1>IIF(m.d_u2<{01.04.2017}, 90, 30)
   LOOP
  ENDIF 
  
  INSERT INTO outfl FROM MEMVAR 
  
 ENDSCAN 
 SET RELATION OFF INTO cur_tmpl
 USE IN cur_tmpl
 
 SELECT outfl

 PUBLIC oExcel AS Excel.Application
 WAIT "������ MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 WAIT "������������ ������..." WINDOW NOWAIT 

 oExcel.UseSystemSeparators = .F.
 oExcel.DecimalSeparator = '.'

 oexcel.ReferenceStyle= -4150  && xlR1C1
 
 oexcel.SheetsInNewWorkbook = 1
 oBook = oExcel.WorkBooks.Add
 
 oSheet = oBook.WorkSheets(1)
 oSheet.Select
 oSheet.Name = '�������'

 =HeadOfTheList()

 nRow = 8
 nRnn = 1
 SCAN 

  m.mcod = mcod 
  m.lpuname = lpuname
  IF !SEEK(m.mcod, 'curmcod')
   INSERT INTO curmcod (mcod, lpuname) VALUES (m.mcod, m.lpuname)
  ENDIF 
 
  oExcel.Cells(nRow,01).HorizontalAlignment = -4131
  oExcel.Cells(nRow,02).HorizontalAlignment = -4131
  oExcel.Cells(nRow,03).HorizontalAlignment = -4131
  oExcel.Cells(nRow,04).HorizontalAlignment = -4131
  oExcel.Cells(nRow,13).HorizontalAlignment = -4131
  oExcel.Cells(nRow,14).HorizontalAlignment = -4131
  oExcel.Cells(nRow,15).HorizontalAlignment = -4131
  oExcel.Cells(nRow,16).HorizontalAlignment = -4131
  oExcel.Cells(nRow,17).HorizontalAlignment = -4131
  oExcel.Cells(nRow,18).HorizontalAlignment = -4131
  oExcel.Cells(nRow,19).HorizontalAlignment = -4131
  oExcel.Cells(nRow,20).HorizontalAlignment = -4131
  oExcel.Cells(nRow,21).HorizontalAlignment = -4131

  oExcel.Rows(nRow).RowHeight=20
  oExcel.Rows(nRow).VerticalAlignment = -4108

  oExcel.Cells(nRow,1).Value  = nRnn
  oExcel.Cells(nRow,2).Value  = mcod
  oExcel.Cells(nRow,3).Value  = fam
  oExcel.Cells(nRow,4).Value  = im
  oExcel.Cells(nRow,5).Value  = ot
  oExcel.Cells(nRow,6).Value  = DTOC(dr)
  oExcel.Cells(nRow,7).Value  = c_i
  oExcel.Cells(nRow,8).Value  = sn_pol
  oExcel.Cells(nRow,9).Value  = DTOC(d_u)
  oExcel.Cells(nRow,10).Value = k_u
  oExcel.Cells(nRow,11).Value = ds
  oExcel.Cells(nRow,12).Value = otd
  oExcel.Cells(nRow,13).Value = pcod
  oExcel.Cells(nRow,14).Value = PADL(cod,6,'0')
  oExcel.Cells(nRow,15).Value = DTOC(d_beg)
  oExcel.Cells(nRow,16).Value = DTOC(d_end)
  oExcel.Cells(nRow,17).Value = STR(k_u,3)
  oExcel.Cells(nRow,18).Value = STR(n_kd,3)
  oExcel.Cells(nRow,19).Value = Tip
  oExcel.Cells(nRow,20).Value = TRANSFORM(s_all,'999999.99')
  oExcel.Cells(nRow,21).Value = IIF(ischked=.t., '���������','�� ���������')
  
  nRnn = nRnn + 1
  nRow = nRow + 1
 ENDSCAN 

 SELECT curmcod
 
 IF RECCOUNT()>0
  SCAN 
   oSheet = oBook.WorkSheets.Add(,oexcel.ActiveSheet)
   oSheet.Name = mcod
   
   =HeadOfTheList(mcod, lpuname)

   =BodyOfTheList(mcod)

*   nRow = 8
*   nRnn = 1
  ENDSCAN 
 ENDIF 

 USE IN curmcod
 USE IN outfl
 USE IN curgosps

 WAIT CLEAR 

 BookName = 'dblgosps'+m.qcod+PADL(DAY(DATE()),2,'0')+PADL(MONTH(DATE()),2,'0')
 IF fso.FileExists(pmee+'\'+BookName+'.xls')
  fso.DeleteFile(pmee+'\'+BookName+'.xls')
 ENDIF 

 oBook.SaveAs(pmee+'\'+BookName,18)
 oExcel.Visible = .T.


RETURN 

FUNCTION seldblgosp(m.lpath)
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
  IF !IsGosp(m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\nsi\sprlpuxx.dbf')
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
  IF OpenFile(m.llcpath+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
   IF USED('sprlpu')
    USE IN sprlpu
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
   IF USED('merror')
    USE IN merror
   ENDIF 
   IF USED('sprlpu')
    USE IN sprlpu
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
  IF OpenFile(m.llcpath+'\nsi'+'\tarifn', 'tarif', 'shar', 'cod')>0
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   IF USED('merror')
    USE IN merror
   ENDIF 
   IF USED('sprlpu')
    USE IN sprlpu
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
   IF EMPTY(tip)
    LOOP 
   ENDIF 

   m.sn_pol = sn_pol
   m.c_i    = c_i
   m.fam    = people.fam
   m.im     = people.im
   m.ot     = people.ot
   m.dr     = people.dr 
   m.w      = people.w
*   m.ul     = people.ul
*   m.dom    = people.dom
*   m.kor    = people.kor
*   m.str    = people.str
*   m.kv     = people.kv
   m.d_u    = d_u
   m.cod    = cod
   m.ds     = ds
   m.dss    = LEFT(ds,3)
   m.k_u    = k_u 
   m.otd    = otd
   m.pcod   = pcod
   
   m.d_beg  = people.d_beg
   m.d_end  = people.d_end
   m.cod    = cod
   m.s_all  = s_all
   m.ischked = IIF(EMPTY(merror.recid), .f., .t.)
   m.tip = tip

   m.codname = IIF(SEEK(m.cod, 'tarif'), tarif.comment, '')

   m.n_kd = tarif.n_kd

   IF !SEEK(m.c_i, 'curgosps')
*    INSERT INTO curgosps (period,lpuid,mcod,sn_pol,c_i,fam,im,ot,dr,w,ul,dom,kor,str,kv,d_u,k_u,ds,dss,otd,pcod,lpuname) VALUES ;
     (m.gcperiod,m.lpuid,m.mcod,m.sn_pol,m.c_i,m.fam,m.im,m.ot,m.dr,m.w,m.ul,m.dom,m.kor,m.str,m.kv,m.d_u,m.k_u,m.ds,m.dss,m.otd,m.pcod,m.lpuname) 
    INSERT INTO curgosps (period,lpuid,mcod,sn_pol,c_i,fam,im,ot,dr,w,d_u,k_u,n_kd,ds,dss,otd,pcod,d_beg,d_end,cod,s_all,ischked,tip,lpuname) VALUES ;
     (m.gcperiod,m.lpuid,m.mcod,m.sn_pol,m.c_i,m.fam,m.im,m.ot,m.dr,m.w,m.d_u,m.k_u,m.n_kd,m.ds,m.dss,m.otd,m.pcod,m.d_beg,m.d_end,m.cod,m.s_all,m.ischked,m.tip,m.lpuname) 
   ELSE 
     m.ok_u = curgosps.k_u
     m.nk_u = m.ok_u + m.k_u
    IF m.d_u > curgosps.d_u
     UPDATE curgosps SET d_u=m.d_u, ds=m.ds, dss=m.dss, k_u=m.nk_u WHERE c_i=m.c_i
    ELSE 
     UPDATE curgosps SET k_u=m.nk_u WHERE c_i=m.c_i
    ENDIF 
   ENDIF 

  ENDSCAN 
  SET RELATION OFF INTO error
  SET RELATION OFF INTO people
  SET RELATION OFF INTO merror
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
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
 
  SELECT aisoms

 ENDSCAN 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 

RETURN 

FUNCTION IsGosp(lcmcod)
 m.lnlputip = INT(VAL(SUBSTR(lcmcod,3,2)))
RETURN IIF(BETWEEN(m.lnlputip,40,67), .t., .f.)

FUNCTION HeadOfTheList(llcmcod, llcname)
 oexcel.Cells.Font.Name='Calibri'
 oexcel.ActiveSheet.PageSetup.Orientation=2
 oExcel.Cells.NumberFormat = '@'


 oRange = oExcel.Range(oExcel.Cells(1,1), oExcel.Cells(1,12))
 oRange.Merge
 oExcel.Cells(1,1).Value='��������� �������������� '
 oExcel.Cells(1,1).HorizontalAlignment = -4108
 oExcel.Cells(1,1).Font.Size = 12
 oExcel.Cells(1,1).Font.Bold = .F.
 oExcel.Cells(1,1).Font.Italic = .T.

 oRange = oExcel.Range(oExcel.Cells(3,1), oExcel.Cells(3,12))
 oRange.Merge
 oExcel.Cells(3,1).Value='��� '+ALLTRIM(m.qname)
 oExcel.Cells(3,1).HorizontalAlignment = -4108
 oExcel.Cells(3,1).Font.Size = 12
 oExcel.Cells(3,1).Font.Italic = .T.

 oRange = oExcel.Range(oExcel.Cells(5,1), oExcel.Cells(5,12))
 oRange.Merge
 oExcel.Cells(5,1).Value='���� '+PADL(DAY(DATE()),2,'0')+' '+LOWER(NameOfMonth2(MONTH(DATE())))+;
  ' '+STR(YEAR(DATE()),4)+' ����'
 oExcel.Cells(5,1).HorizontalAlignment = -4108
 oExcel.Cells(5,1).Font.Size = 12
 oExcel.Cells(5,1).Font.Italic = .T.
 
 oExcel.Columns(1).ColumnWidth = 3
 oExcel.Columns(2).ColumnWidth = 7
 oExcel.Columns(3).ColumnWidth = 15
 oExcel.Columns(4).ColumnWidth = 15
 oExcel.Columns(5).ColumnWidth = 15
 oExcel.Columns(6).ColumnWidth = 9
 oExcel.Columns(7).ColumnWidth = 15
 oExcel.Columns(8).ColumnWidth = 17
 oExcel.Columns(9).ColumnWidth = 9
 oExcel.Columns(10).ColumnWidth = 5
 oExcel.Columns(11).ColumnWidth = 7
 oExcel.Columns(12).ColumnWidth = 5
 oExcel.Columns(13).ColumnWidth = 12
 
 oExcel.Cells(7,1).Value  = '� �/�'
 oExcel.Cells(7,2).Value  = '��� ���'
 oExcel.Cells(7,3).Value  = '�������'
 oExcel.Cells(7,4).Value  = '���'
 oExcel.Cells(7,5).Value  = '��������'
 oExcel.Cells(7,6).Value  = '���� ��������'
 oExcel.Cells(7,7).Value  = '�����'
 oExcel.Cells(7,8).Value  = '�����'
 oExcel.Cells(7,9).Value  = '����'
 oExcel.Cells(7,10).Value = '�/�'
 oExcel.Cells(7,11).Value = '�������'
 oExcel.Cells(7,12).Value = '���������'
 oExcel.Cells(7,13).Value = '����'
 oExcel.Cells(7,14).Value = '���'
 oExcel.Cells(7,15).Value = '����'
 oExcel.Cells(7,16).Value = '����'
 oExcel.Cells(7,17).Value = '���-�� ����'
 oExcel.Cells(7,18).Value = '���-�� ����'
 oExcel.Cells(7,19).Value = '���'
 oExcel.Cells(7,20).Value = '�����'
 oExcel.Cells(7,21).Value = '���������?'
 oExcel.Rows(7).RowHeight=40
 oExcel.Rows(7).VerticalAlignment = -4108
RETURN 

FUNCTION BodyOfTheList(lcmcod)
 nRow = 8
 nRnn = 1

 SELECT outfl
 SCAN
  
  IF mcod!=lcmcod
   LOOP 
  ENDIF 
 
  oExcel.Cells(nRow,01).HorizontalAlignment = -4131
  oExcel.Cells(nRow,02).HorizontalAlignment = -4131
  oExcel.Cells(nRow,03).HorizontalAlignment = -4131
  oExcel.Cells(nRow,04).HorizontalAlignment = -4131
  oExcel.Cells(nRow,13).HorizontalAlignment = -4131
  oExcel.Cells(nRow,14).HorizontalAlignment = -4131
  oExcel.Cells(nRow,15).HorizontalAlignment = -4131
  oExcel.Cells(nRow,16).HorizontalAlignment = -4131
  oExcel.Cells(nRow,17).HorizontalAlignment = -4131
  oExcel.Cells(nRow,18).HorizontalAlignment = -4131
  oExcel.Cells(nRow,19).HorizontalAlignment = -4131
  oExcel.Cells(nRow,20).HorizontalAlignment = -4131
  oExcel.Cells(nRow,21).HorizontalAlignment = -4131

  oExcel.Rows(nRow).RowHeight=20
  oExcel.Rows(nRow).VerticalAlignment = -4108

  oExcel.Cells(nRow,1).Value  = nRnn
  oExcel.Cells(nRow,2).Value  = mcod
  oExcel.Cells(nRow,3).Value  = fam
  oExcel.Cells(nRow,4).Value  = im
  oExcel.Cells(nRow,5).Value  = ot
  oExcel.Cells(nRow,6).Value  = DTOC(dr)
  oExcel.Cells(nRow,7).Value  = c_i
  oExcel.Cells(nRow,8).Value  = sn_pol
  oExcel.Cells(nRow,9).Value  = DTOC(d_u)
  oExcel.Cells(nRow,10).Value = k_u
  oExcel.Cells(nRow,11).Value = ds
  oExcel.Cells(nRow,12).Value = otd
  oExcel.Cells(nRow,13).Value = pcod
  oExcel.Cells(nRow,14).Value = PADL(cod,6,'0')
  oExcel.Cells(nRow,15).Value = DTOC(d_beg)
  oExcel.Cells(nRow,16).Value = DTOC(d_end)
  oExcel.Cells(nRow,17).Value = STR(k_u,3)
  oExcel.Cells(nRow,18).Value = STR(n_kd,3)
  oExcel.Cells(nRow,19).Value = Tip
  oExcel.Cells(nRow,20).Value = TRANSFORM(s_all,'999999.99')
  oExcel.Cells(nRow,21).Value = IIF(ischked=.t., '���������','�� ���������')
  
  nRnn = nRnn + 1
  nRow = nRow + 1
 ENDSCAN 

 SELECT curmcod
RETURN 
