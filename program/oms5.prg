FUNCTION  oms5(lpucod, IsVisible, IsQuit)

# DEFINE xlDiagonalDown	     5	&& ������������ �� �������� ������ ���� � ������ ������ ������ ������ � ���������
# DEFINE xlDiagonalUp	     6	&& ������������ �� ������� ������ ���� � ������ ������� ������ ������ � ���������.
# DEFINE xlEdgeBottom	     9	&& ������� ��� ����� ��������� �����
# DEFINE xlEdgeLeft	         7  && ����� ��� ����� ��������� �����.
# DEFINE xlEdgeRight	    10	&& ������ ��� ����� ��������� �����.
# DEFINE xlEdgeTop	         8	&& ������� ��� ����� ��������� �����.
# DEFINE xlInsideHorizontal	12	&& �������������� ������� ���� ���������� ����� ���������
# DEFINE xlInsideVertical	11	&& ������������ ������� ���� ���������� ����� ���������

# DEFINE xlContinuous	     1	&& ����������� �����
# DEFINE xlDash	         -4115	&& ���������� �����
# DEFINE xlDashDot	         4	&& ������� � ������
# DEFINE xlDashDotDot	     5	&& ������� � ����� ������� ������ �������
# DEFINE xlDot	         -4118	&& ����� �� �����
# DEFINE xlDouble	     -4119	&& ������� �����
# DEFINE xlLineStyleNone -4142	&& ��� �����
# DEFINE xlSlantDashDot	    13	&& ��������� ����������

# DEFINE xlHairline	    1	&& ����� ������ �������
# DEFINE xlMedium	-4138	&& ������� �������
# DEFINE xlThick	    4	&& ������� �������
# DEFINE xlThin	        2   && ������ �������

m.mcod = lpucod

pMail = fso.GetParentFolderName(pbin)+'\MEE'

IF !fso.FolderExists(pMail)
 fso.CreateFolder(pMail)
ENDIF 

m.SortTip = '0'
oal = ALIAS()
IF fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\outs.dbf')
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\outs', 'outs', 'shar')>0
  IF USED('outs')
   USE IN outs
  ENDIF 
 ENDIF 
ENDIF 
CREATE CURSOR TipSort (name c(20), cod c(1))
INSERT INTO TipSort (name,cod) VALUES ('�� �����������','0')
INSERT INTO TipSort (name,cod) VALUES ('���� ������/���','1')
INSERT INTO TipSort (name,cod) VALUES ('��� ������/���','2')
INSERT INTO TipSort (name,cod) VALUES ('�������','3')
SELECT (oal)

DO FORM SortOms5

BookName = pMail+'\'+m.mcod+'_'+ALLTRIM(sn_pol)
m.IsOpDoc = IsOpenExcelDoc(m.mcod+'_'+ALLTRIM(sn_pol))
IF m.IsOpDoc
 IF !CloseExcelDoc(m.mcod+'_'+ALLTRIM(sn_pol))
  MESSAGEBOX('���� '+m.mcod+'_'+ALLTRIM(sn_pol)+' ������!',0+64,'')
  RETURN .f. 
 ENDIF 
ENDIF 

PUBLIC oExcel AS Excel.Application
WAIT "������ MS Excel..." WINDOW NOWAIT 
TRY 
 oExcel=GETOBJECT(,"Excel.Application")
CATCH 
 oExcel=CREATEOBJECT("Excel.Application")
ENDTRY 
WAIT CLEAR 

oexcel.UseSystemSeparators= .F.
oexcel.DecimalSeparator = '.'

oexcel.ReferenceStyle= -4150  && xlR1C1
 
oexcel.SheetsInNewWorkbook=1
oBook = oExcel.WorkBooks.Add
oexcel.Cells.Font.Name='Calibri'
oexcel.ActiveSheet.PageSetup.Orientation=2


oSheet = oBook.WorkSheets(1)
oSheet.Select
 
FOR iii=1 TO 12
 oexcel.Columns(iii).NumberFormat='@'
ENDFOR 

nCell = 1
orec = RECNO()

LpuName  = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname),'')
CokrCod  = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.cokr), '')
CokrName = IIF(SEEK(m.cokrcod, 'admokr'), admokr.name_okr, '')

m.prmcod = people.prmcod
m.lpupr = IIF(SEEK(m.prmcod, 'sprlpu'), ALLTRIM(sprlpu.fullname)+', '+m.prmcod,'')
m.sppr  = ''

m.tipp = IIF(FIELD('tipp', 'people')='TIPP', people.tipp, '')
m.ppolis = ''

IF USED('outs')
 DO CASE 
  CASE m.tipp = '�'
   m.ppolis = SUBSTR(people.sn_pol,7,9)
   m.sppr = IIF(SEEK(m.ppolis, 'outs', 'vsn'), IIF(outs.spos_tera=1,'��������������� (1)','�� ��������� (2)')+', '+DTOC(outs.date_tera), '')
  CASE m.tipp = '�'
   m.ppolis = LEFT(people.sn_pol,16)
   m.sppr = IIF(SEEK(m.ppolis, 'outs', 'kms'), IIF(outs.spos_tera=1,'��������������� (1)','�� ��������� (2)')+', '+DTOC(outs.date_tera), '')
  CASE m.tipp = '�'
   m.ppolis = people.enp
   m.sppr = IIF(SEEK(m.ppolis, 'outs', 'enp'), IIF(outs.spos_tera=1,'��������������� (1)','�� ��������� (2)')+', '+DTOC(outs.date_tera), '')
  OTHERWISE 
   m.ppolis = LEFT(people.sn_pol,16)
   m.sppr = IIF(SEEK(m.ppolis, 'outs', 'kms'), IIF(outs.spos_tera=1,'��������������� (1)','�� ��������� (2)')+', '+DTOC(outs.date_tera), '')
 ENDCASE 
 USE IN outs 
ENDIF 

 WITH oExcel.Sheets(1)
  .cells(1,1).Value2 = '���: ' + m.lpuname + ', ' + m.cokrname + ', ' + m.mcod
  .cells(1,1).Characters(1,4).Font.Bold = .t.
  .cells(1,1).Characters(1,4).Font.Italic = .t.

  .cells(2,1).Value2 = '���: ' + m.qname
  .cells(2,1).Characters(1,4).Font.Bold = .t.
  .cells(2,1).Characters(1,4).Font.Italic = .t.

  .cells(4,1).Value2 = '�������: ' + ALLTRIM(people.fam)+' '+ALLTRIM(people.im)+' '+ALLTRIM(people.ot)+', '+DTOC(people.dr)
  .cells(4,1).Characters(1,8).Font.Bold = .t.
  .cells(4,1).Characters(1,8).Font.Italic = .t.

  .cells(5,1).Value2 = '�����: ' + ALLTRIM(sn_pol)
  .cells(5,1).Characters(1,6).Font.Bold = .t.
  .cells(5,1).Characters(1,6).Font.Italic = .t.

  .cells(7,1).Value2 = '�����: ' + ALLTRIM(talon.c_i)
  .cells(7,1).Characters(1,6).Font.Bold = .t.
  .cells(7,1).Characters(1,6).Font.Italic = .t.

  .cells(8,1).Value2 = '��� �����������: ' + m.lpupr
  .cells(8,1).Characters(1,16).Font.Bold = .t.
  .cells(8,1).Characters(1,16).Font.Italic = .t.

  .cells(9,1).Value2 = '������ ������������: ' + m.sppr
  .cells(9,1).Characters(1,20).Font.Bold = .t.
  .cells(9,1).Characters(1,20).Font.Italic = .t.

  .cells(10,1).Value2 = '���� �� ��������� ����������� ������ �� ���������� ��������� ��������� ���'
  .cells(11,3).Value2 = '�� '+ NameOfMonth(tMonth)+ ' '+STR(tYear,4)+' ����'
  .cells(10,1).Font.Size = 11
  .cells(11,1).Font.Size = 11
  .cells(10,1).Font.Bold = .T.
  .cells(11,1).Font.Bold = .T.
  .cells(10,1).HorizontalAlignment=-4108
  .cells(11,3).HorizontalAlignment=-4108

  FOR nRow=1 TO 11
   oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,11))
   oRange.Merge
  ENDFOR  
 
  .cells(13,1).Value2 = '����'
  .cells(13,2).Value2 = '������'
  .cells(13,3).Value2 = '���'
  .cells(13,4).Value2 = '�������'
  .cells(13,5).Value2 = '������������ ������'
  .cells(13,6).Value2 = '���-��'
  .cells(13,7).Value2 = '�����'
  .cells(13,8).Value2 = '��'
  .cells(13,9).Value2 = '��'
  .cells(13,10).Value2 = '���� ����.'
  .cells(13,11).Value2 = '��� ����.'
  .cells(13,12).Value2 = '��� ����.'
  
  oal = ALIAS()
  m.polis = sn_pol
  SELECT talon 
  oord = ORDER('talon')
  DO CASE 
   CASE m.SortTip = '0'
    SET ORDER TO 
   CASE m.SortTip = '1'
    SET ORDER TO d_u
   CASE m.SortTip = '2'
    SET ORDER TO cod
   CASE m.SortTip = '3'
    SET ORDER TO ds
  ENDCASE 

  nCell = 13
  m.ttlkol = 0
  m.ttlsum = 0
  SCAN 
   IF sn_pol = m.polis
    m.cod = cod 
    nCell = nCell + 1
    .cells(nCell,1).Value2 = DTOC(d_u)
    .cells(nCell,2).Value2 = PADL(cod,6,'0')
    .cells(nCell,3).Value2 = tip
    .cells(nCell,4).Value2 = ds
    .cells(nCell,5).Value2 = IIF(SEEK(m.cod, 'tarif'), tarif.comment, '')
    .cells(nCell,6).Value2 = STR(k_u,3)
    .cells(nCell,7).Value2 = TRANSFORM(s_all, '99 999 999.99')
    .cells(nCell,8).Value2 = d_type
    IF IsUsl(cod) AND FIELD('lpu_ord','talon')='LPU_ORD'
     m.llpuid = lpu_ord
     m.mmcod = IIF(SEEK(m.llpuid, 'sprlpu', 'lpu_id'), sprlpu.mcod, '')
     .cells(nCell,9).Value2 = STR(ord,1)
     .cells(nCell,10).Value2 = DTOC(date_ord)
     .cells(nCell,11).Value2 = m.mmcod
    ENDIF 
    .cells(nCell,12).Value2 = mcod
    m.ttlkol = m.ttlkol + k_u
    m.ttlsum = m.ttlsum + s_all
   ENDIF 
  ENDSCAN 
 .cells(nCell+1,6).Value2 = STR(m.ttlkol,3)
 .cells(nCell+1,7).Value2 = TRANSFORM(m.ttlsum, '99 999 999.99')
  SET ORDER TO (oord)
  SELECT (oal)
 ENDWITH 

GO (orec)

FOR iii=1 TO 12
 oexcel.Columns(iii).AutoFit
ENDFOR 

m.ttlsumsay = '�����: '+cpr(INT(m.ttlsum))+PADL(INT((m.ttlsum-INT(m.ttlsum))*100),2,'0')+' ���.'
oExcel.Sheets(1).cells(nCell+2,1).Value2 = m.ttlsumsay
oExcel.Range(oExcel.Sheets(1).cells(nCell+2,1), oExcel.Sheets(1).cells(nCell+2,8)).Merge 

IF fso.FileExists(pMail+'\'+m.mcod+'_'+ALLTRIM(sn_pol)+'.xls')
 fso.DeleteFile(pMail+'\'+m.mcod+'_'+ALLTRIM(sn_pol)+'.xls')
ENDIF 

oBook.SaveAs(BookName,18)

IF IsVisible == .T. 
 oExcel.Visible = .T.
ELSE 
 oBook.Close(0)
 IF IsQuit
  oExcel.Quit
 ENDIF 
ENDIF 

RETURN 
