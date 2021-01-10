PROCEDURE Sel50pers
 IF MESSAGEBOX(CHR(13)+CHR(10)+'�������� �������� (<50%)'+CHR(13)+CHR(10)+;
  '� ������� (50%) ��������������?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
* CREATE CURSOR curdeads (period c(7), lpuid i(4), mcod c(7), sn_pol c(25), c_i c(30), ;
  fam c(25), im c(25), ot c(25), dr d, w n(1), ul n(5), dom c(7), ;
  kor c(5), str c(5), kv c(5), d_u d, ds c(6), otd c(4), pcod c(10), ;
  cod n(6), k_u n(3), tip c(1), n_kd n(3),s_all n(11,2), d_beg d, d_end d, codname c(40), lpuname c(40), ischked l)
 CREATE CURSOR curdeads (period c(7), lpuid i(4), mcod c(7), sn_pol c(25), c_i c(30), ;
  fam c(25), im c(25), ot c(25), dr d, w n(1), d_u d, ds c(6), otd c(4), pcod c(10), ;
  cod n(6), k_u n(3), tip c(1), n_kd n(3),s_all n(11,2), d_beg d, d_end d, codname c(40), lpuname c(40), ischked l)

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
  =seldeadsone(m.lpath)
  WAIT CLEAR 

 NEXT 

 CREATE CURSOR curmcod (mcod c(7), lpuname c(40))
 INDEX on mcod TAG mcod
 SET ORDER TO mcod 

 SELECT curdeads
 outfile = pmee+'\sldeads'
 
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
  oExcel.Cells(nRow,7).Value  = sn_pol
  oExcel.Cells(nRow,8).Value  = '������� ��������'
  oExcel.Cells(nRow,9).Value  = pcod
  oExcel.Cells(nRow,10).Value = c_i
  oExcel.Cells(nRow,11).Value = otd
  oExcel.Cells(nRow,12).Value = PADL(cod,6,'0')
  oExcel.Cells(nRow,13).Value = codname
  oExcel.Cells(nRow,14).Value = ds
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
 COPY TO &outfile
* USE 
 
 SELECT curmcod
 IF RECCOUNT()>0
  SCAN 
   oSheet = oBook.WorkSheets.Add(,oexcel.ActiveSheet)
   oSheet.Name = mcod
   
   =HeadOfTheList(mcod, lpuname)
   =BodyOfTheList(mcod)

 nRow = 8
 nRnn = 1
  ENDSCAN 
 ENDIF 
 USE IN curmcod
 USE IN curdeads

 WAIT CLEAR 

 BookName = 'LongAndShortHosps'+m.qcod+PADL(DAY(DATE()),2,'0')+PADL(MONTH(DATE()),2,'0')
 IF fso.FileExists(pmee+'\'+BookName+'.xls')
  fso.DeleteFile(pmee+'\'+BookName+'.xls')
 ENDIF 

 oBook.SaveAs(pmee+'\'+BookName,18)
 oExcel.Visible = .T.

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
   m.k_u = k_u
   m.n_kd = tarif.n_kd
   m.koeff = ROUND(m.k_u/m.n_kd,0)
   
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
*   m.ul     = people.ul
*   m.dom    = people.dom
*   m.kor    = people.kor
*   m.str    = people.str
*   m.kv     = people.kv
   m.d_beg  = people.d_beg
   m.d_end  = people.d_end
   m.d_u    = d_u
   m.ds     = ds
   m.otd    = otd
   m.pcod   = pcod
   m.cod    = cod
   m.k_u    = k_u
   m.s_all  = s_all
   m.ischked = IIF(EMPTY(merror.recid), .f., .t.)
   m.tip = tip

   m.codname = IIF(SEEK(m.cod, 'tarif'), tarif.comment, '')

*   INSERT INTO curdeads (period,lpuid,mcod,sn_pol,c_i,fam,im,ot,dr,w,;
    ul,dom,kor,str,kv,d_u,ds,otd,pcod, cod,k_u, n_kd,s_all, d_beg, d_end, codname,lpuname,ischked,tip) VALUES ;
    (m.lcperiod,m.lpuid,m.mcod,m.sn_pol,m.c_i,m.fam,m.im,m.ot,m.dr,m.w,;
     m.ul,m.dom,m.kor,m.str,m.kv,m.d_u,m.ds,m.otd,m.pcod,m.cod,m.k_u, m.n_kd,m.s_all, ;
     m.d_beg, m.d_end, m.codname,m.lpuname,m.ischked,m.tip) 
   INSERT INTO curdeads (period,lpuid,mcod,sn_pol,c_i,fam,im,ot,dr,w,;
    d_u,ds,otd,pcod, cod,k_u, n_kd,s_all, d_beg, d_end, codname,lpuname,ischked,tip) VALUES ;
    (m.lcperiod,m.lpuid,m.mcod,m.sn_pol,m.c_i,m.fam,m.im,m.ot,m.dr,m.w,;
     m.d_u,m.ds,m.otd,m.pcod,m.cod,m.k_u, m.n_kd,m.s_all, ;
     m.d_beg, m.d_end, m.codname,m.lpuname,m.ischked,m.tip) 

  ENDSCAN 
  SET RELATION OFF INTO merror
  SET RELATION OFF INTO error
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

 ENDSCAN 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 

RETURN 

FUNCTION HeadOfTheList(llcmcod, llcname)
 oexcel.Cells.Font.Name='Calibri'
 oexcel.ActiveSheet.PageSetup.Orientation=2
 oExcel.Cells.NumberFormat = '@'

 oRange = oExcel.Range(oExcel.Cells(1,1), oExcel.Cells(1,20))
 oRange.Merge
 oExcel.Cells(1,1).Value='������ �������������� � ����������� ����� ��� �� 50% ��� ���������� ����� ��� �� 50% �� ����������� ����������������� �������'
 oExcel.Cells(1,1).HorizontalAlignment = -4108
 oExcel.Cells(1,1).Font.Size = 12
 oExcel.Cells(1,1).Font.Bold = .F.
 oExcel.Cells(1,1).Font.Italic = .T.
 oExcel.Rows(1).RowHeight = 30
 oExcel.Rows(1).VerticalAlignment = -4108
 
 oRange = oExcel.Range(oExcel.Cells(2,1), oExcel.Cells(2,2))
 oRange.Merge
 oExcel.Cells(2,1).Value = '�� ���:'
 oRange = oExcel.Range(oExcel.Cells(2,3), oExcel.Cells(2,20))
 oRange.Merge
 oExcel.Cells(2,3).Value = IIF(EMPTY(llcmcod), '������� �� ���� ���', ALLTRIM(llcname))
 oExcel.Cells(2,3).Font.Size = 12
 oExcel.Cells(2,3).Font.Bold = .F.
 oExcel.Cells(2,3).Font.Italic = .T.
 oExcel.Rows(2).RowHeight = 30
 oExcel.Rows(2).VerticalAlignment = -4108

 oRange = oExcel.Range(oExcel.Cells(3,1), oExcel.Cells(3,3))
 oRange.Merge
 oExcel.Cells(3,1).Value = '�� ������:'
 oExcel.Cells(3,4).Value = '�:'
 oExcel.Cells(3,5).Value = '��:'
 
 oExcel.Columns(1).ColumnWidth = 3
 oExcel.Columns(2).ColumnWidth = 7
 oExcel.Columns(3).ColumnWidth = 15
 oExcel.Columns(4).ColumnWidth = 15
 oExcel.Columns(5).ColumnWidth = 15
 oExcel.Columns(6).ColumnWidth = 9
 oExcel.Columns(7).ColumnWidth = 17
 oExcel.Columns(8).ColumnWidth = 5
 oExcel.Columns(9).ColumnWidth = 12
 oExcel.Columns(10).ColumnWidth = 17
 oExcel.Columns(11).ColumnWidth = 9
 oExcel.Columns(12).ColumnWidth = 10
 oExcel.Columns(13).ColumnWidth = 50
 oExcel.Columns(14).ColumnWidth = 10
 oExcel.Columns(15).ColumnWidth = 9
 oExcel.Columns(16).ColumnWidth = 9
 oExcel.Columns(17).ColumnWidth = 5
 oExcel.Columns(18).ColumnWidth = 5
 oExcel.Columns(19).ColumnWidth = 5
 
 oExcel.Cells(7,1).Value  = '� �/�'
 oExcel.Cells(7,2).Value  = '��� ���'
 oExcel.Cells(7,3).Value  = '�������'
 oExcel.Cells(7,4).Value  = '���'
 oExcel.Cells(7,5).Value  = '��������'
 oExcel.Cells(7,6).Value  = '���� ��������'
 oExcel.Cells(7,7).Value  = '�����'
 oExcel.Cells(7,8).Value  = '������� ��������'
 oExcel.Cells(7,9).Value  = '����'
 oExcel.Cells(7,10).Value = '�����'
 oExcel.Cells(7,11).Value = '���������'
 oExcel.Cells(7,12).Value = '������/���'
 oExcel.Cells(7,13).Value = '������������ ������/����'
 oExcel.Cells(7,14).Value = '�������'
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

 SELECT curdeads
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

  oExcel.Rows(nRow).RowHeight=20
  oExcel.Rows(nRow).VerticalAlignment = -4108

  oExcel.Cells(nRow,1).Value  = nRnn
  oExcel.Cells(nRow,2).Value  = mcod
  oExcel.Cells(nRow,3).Value  = fam
  oExcel.Cells(nRow,4).Value  = im
  oExcel.Cells(nRow,5).Value  = ot
  oExcel.Cells(nRow,6).Value  = DTOC(dr)
  oExcel.Cells(nRow,7).Value  = sn_pol
  oExcel.Cells(nRow,8).Value  = '������� ��������'
  oExcel.Cells(nRow,9).Value  = pcod
  oExcel.Cells(nRow,10).Value = c_i
  oExcel.Cells(nRow,11).Value = otd
  oExcel.Cells(nRow,12).Value = PADL(cod,6,'0')
  oExcel.Cells(nRow,13).Value = codname
  oExcel.Cells(nRow,14).Value = ds
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