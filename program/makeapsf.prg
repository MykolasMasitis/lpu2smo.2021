PROCEDURE MakeAPSF

IF MESSAGEBOX(CHR(13)+CHR(10)+'�� ������ ������������'+CHR(13)+CHR(10)+;
 '��� �������� ������-������?'+CHR(13)+CHR(10))==7
 RETURN 
ENDIF 

IF OpenFile(pcommon+'\lpudogs', 'lpudogs', 'shar', 'mcod')>0
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 RETURN 
ENDIF 
IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\admokrxx', 'admokr', 'shar', 'cokr')>0
 IF USED('admokr')
  USE IN admokr
 ENDIF 
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 RETURN 
ENDIF 

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'cokr')>0
 IF USED('admokr')
  USE IN admokr
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 RETURN 
ENDIF 
IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 IF USED('admokr')
  USE IN admokr
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 RETURN 
ENDIF 

SELECT sprlpu
SET RELATION TO mcod INTO aisoms

IF !fso.FolderExists(pout+'\'+gcperiod)
 fso.CreateFolder(pout+'\'+gcperiod)
ENDIF 

pdir = pout+'\'+gcperiod

PUBLIC oExcel AS Excel.Application

WAIT "������ MS Excel..." WINDOW NOWAIT 
TRY 
 oExcel=GETOBJECT(,"Excel.Application")
CATCH 
 oExcel=CREATEOBJECT("Excel.Application")
ENDTRY 
WAIT CLEAR 

WAIT "������������ ������..." WINDOW NOWAIT 

oexcel.UseSystemSeparators = .F.
oexcel.DecimalSeparator = '.'

oexcel.ReferenceStyle= -4150  && xlR1C1
 
oexcel.SheetsInNewWorkbook=1
oBook = oExcel.WorkBooks.Add
oexcel.Cells.Font.Name='Calibri'
oexcel.Cells.Font.Size=9
oexcel.ActiveSheet.PageSetup.Orientation=2

BookName = pdir+'\ActSF'
oSheet = oBook.WorkSheets(1)
oSheet.Select
 
nCell = 1

WITH oExcel
 .Columns(01).ColumnWidth = 3
 .Columns(02).ColumnWidth = 20
 .Columns(03).ColumnWidth = 10
 .Columns(04).ColumnWidth = 20
 .Columns(05).ColumnWidth = 10
 .Columns(06).ColumnWidth = 07
 .Columns(07).ColumnWidth = 07
 .Columns(08).ColumnWidth = 07
 .Columns(09).ColumnWidth = 07
 .Columns(10).ColumnWidth = 09
 .Columns(11).ColumnWidth = 09
 .Columns(12).ColumnWidth = 10
 .Columns(13).ColumnWidth = 10
 .Columns(14).ColumnWidth = 10

 .Columns(02).NumberFormat = '@'
 .Columns(02).WrapText = .t.
 .Columns(03).NumberFormat = '@'
 .Columns(04).NumberFormat = '@'
 .Columns(05).NumberFormat = "#,##0.00"
 .Columns(06).NumberFormat = "#,##0.00"
 .Columns(07).NumberFormat = "#,##0.00"
 .Columns(08).NumberFormat = "#,##0.00"
 .Columns(09).NumberFormat = "#,##0.00"
 .Columns(10).NumberFormat = "#,##0.00"
 .Columns(11).NumberFormat = "#,##0.00"
 .Columns(12).NumberFormat = "#,##0.00"
 .Columns(13).NumberFormat = "#,##0.00"
 .Columns(14).NumberFormat = "#,##0.00"
ENDWITH 

WITH oExcel.Sheets(1)
 .cells(1,1).Value2 = '��� �������� ������-������'
 .cells(2,1).Value2 = '�� ����������� ������ �� ���������� ��������� ��������� ��� (�� �������� �������� �������c��� ������)'
 .cells(3,1).Value2 = '�� ��� '+m.qname
 .cells(4,1).Value2 = '�� '+ NameOfMonth(tMonth)+ ' '+STR(tYear,4)+' ����'
 .cells(1,1).Font.Size = 11
 .cells(2,1).Font.Size = 11
 .cells(3,1).Font.Size = 11
 .cells(4,1).Font.Size = 11
 .cells(1,1).Font.Bold = .T.
 .cells(2,1).Font.Bold = .T.
 .cells(3,1).Font.Bold = .T.
 .cells(4,1).Font.Bold = .T.
 .cells(1,1).HorizontalAlignment=-4108
 .cells(2,1).HorizontalAlignment=-4108
 .cells(3,1).HorizontalAlignment=-4108
 .cells(4,1).HorizontalAlignment=-4108
 FOR nRow=1 TO 4
  oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,14))
  oRange.Merge
 ENDFOR  
 
ENDWITH 

m.fstring = 7 && ������ ������

WITH oExcel.Sheets(1)
 .Cells(m.fstring,01).Value2 = '� �\�'
 .Cells(m.fstring,02).Value2 = '������������ ��� (������������ ����) � ������� �� �������'
 .Cells(m.fstring,03).Value2 = '����� ������-������, �������������� ���'
 .Cells(m.fstring,04).Value2 = '����� �������� � ���'
 .Cells(m.fstring,05).Value2 = '����� ������-������, ��������-������ ���'
 .Cells(m.fstring,06).Value2 = 'zh-������'
 .Cells(m.fstring,07).Value2 = '���'
 .Cells(m.fstring,08).Value2 = '���'
 .Cells(m.fstring,09).Value2 = '����'
 .Cells(m.fstring,10).Value2 = '�����'
 .Cells(m.fstring,11).Value2 = '���� �� ������ �������'
 .Cells(m.fstring,12).Value2 = '� ����������'
 .Cells(m.fstring,13).Value2 = '� ������'
 .Cells(m.fstring,14).Value2 = '���� �� ����� �������'

 oExcel.Range(oExcel.Cells(m.fstring-1,1), oExcel.Cells(m.fstring,1)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,2), oExcel.Cells(m.fstring,2)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,3), oExcel.Cells(m.fstring,3)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,4), oExcel.Cells(m.fstring,4)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,5), oExcel.Cells(m.fstring,5)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,6), oExcel.Cells(m.fstring,6)).Merge 

 oExcel.Range(oExcel.Cells(m.fstring-1,10), oExcel.Cells(m.fstring,10)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,11), oExcel.Cells(m.fstring,11)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,12), oExcel.Cells(m.fstring,12)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,13), oExcel.Cells(m.fstring,13)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,14), oExcel.Cells(m.fstring,14)).Merge 
 .Cells(m.fstring-1,7).Value2 = '����� ���������� ������-������'
 oExcel.Range(oExcel.Cells(m.fstring-1,7), oExcel.Cells(m.fstring-1,9)).Merge 
ENDWITH 

FOR ncol=1 TO 14
 WITH oExcel.Sheets(1)
  .Cells(m.fstring,ncol).Font.Size = 8
  .Cells(m.fstring,ncol).Font.Bold = .F.
  .Cells(m.fstring,ncol).WrapText = .t.
  .Cells(m.fstring,ncol).HorizontalAlignment = -4108
  .Cells(m.fstring,ncol).VerticalAlignment = -4108
 ENDWITH 
NEXT 

m.fstring = m.fstring + 1
FOR ncolumn=1 TO 14
 WITH oExcel.Sheets(1)
  .cells(m.fstring,ncolumn).Value2 = STR(ncolumn,2)
  .cells(m.fstring,ncolumn).HorizontalAlignment = -4108
 ENDWITH 
NEXT 

oExcel.Range(oExcel.Cells(m.fstring,1), oExcel.Cells(m.fstring,14)).NumberFormat='@'

m.fstring = m.fstring + 1
m.ccokr = cokr
m.nincokr = 1

m.cokrname = IIF(SEEK(m.ccokr, 'admokr'), admokr.name_okr, '')
oExcel.Sheets(1).cells(m.fstring,2).Value2 = m.cokrname
oRange = oExcel.Range(oExcel.Cells(m.fstring,2), oExcel.Cells(m.fstring,14))
oRange.Merge
oRange.HorizontalAlignment = -4108
oRange.Interior.ColorIndex = 40

m.sum04 = 0
m.sum05 = 0
m.sum06 = 0
m.sum07 = 0
m.sum08 = 0
m.sum09 = 0
m.sum10 = 0
m.sum11 = 0
m.sum12 = 0
m.sum13 = 0

m.sum04okr = 0
m.sum05okr = 0
m.sum06okr = 0
m.sum07okr = 0
m.sum08okr = 0
m.sum09okr = 0
m.sum10okr = 0
m.sum11okr = 0
m.sum12okr = 0
m.sum13okr = 0

SCAN FOR aisoms.s_pred>0
 m.fstring = m.fstring + 1
 m.mcod = mcod
 m.numdog = ''
 IF USED('lpudogs')
  m.numdog = IIF(SEEK(m.mcod, 'lpudogs'), lpudogs.dogs, '')
 ENDIF 

 WITH oExcel.Sheets(1)
  .Cells(m.fstring,01).Value2 = m.nincokr
  .Cells(m.fstring,02).Value2 = ALLTRIM(fullname)
  .Cells(m.fstring,03).Value2 = mcod
  .Cells(m.fstring,04).Value2 = m.numdog
  .Cells(m.fstring,05).Value2 = aisoms.s_pred
  .Cells(m.fstring,06).Value2 = 0
  .Cells(m.fstring,07).Value2 = aisoms.sum_flk
  .Cells(m.fstring,08).Value2 = aisoms.e_mee
  .Cells(m.fstring,09).Value2 = aisoms.e_ekmp
  .Cells(m.fstring,10).Value2 = aisoms.s_avans
  .Cells(m.fstring,11).Value2 = aisoms.dolg_b
  .Cells(m.fstring,12).Value2 = aisoms.s_pred - aisoms.sum_flk -  ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b
  .Cells(m.fstring,13).Value2 = IIF(aisoms.s_pred - aisoms.sum_flk -  ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b>0, aisoms.s_pred - aisoms.sum_flk - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b, 0)
  .Cells(m.fstring,14).Value2 = IIF(aisoms.s_pred - aisoms.sum_flk - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b<0, -(aisoms.s_pred - aisoms.sum_flk - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b), 0)
 ENDWITH 

 m.sum04okr = m.sum04okr + aisoms.s_pred
 m.sum05okr = m.sum05okr
 m.sum06okr = m.sum06okr + aisoms.sum_flk
 m.sum07okr = m.sum07okr + aisoms.e_mee
 m.sum08okr = m.sum08okr + aisoms.e_ekmp
 m.sum09okr = m.sum09okr + aisoms.s_avans
 m.sum10okr = m.sum10okr + aisoms.dolg_b
 m.sum11okr = m.sum11okr + (aisoms.s_pred - aisoms.sum_flk - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b)
 m.sum12okr = m.sum12okr + IIF(aisoms.s_pred - aisoms.sum_flk - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b>0, aisoms.s_pred - aisoms.sum_flk - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b, 0)
 m.sum13okr = m.sum13okr + IIF(aisoms.s_pred - aisoms.sum_flk - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b<0, -(aisoms.s_pred - aisoms.sum_flk - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b), 0)

 m.sum04 = m.sum04 + aisoms.s_pred
 m.sum05 = m.sum05
 m.sum06 = m.sum06 + aisoms.sum_flk
 m.sum07 = m.sum07 + aisoms.e_mee
 m.sum08 = m.sum08 + aisoms.e_ekmp
 m.sum09 = m.sum09 + aisoms.s_avans
 m.sum10 = m.sum10 + aisoms.dolg_b
 m.sum11 = m.sum11 + aisoms.s_pred - aisoms.sum_flk - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b
 m.sum12 = m.sum12 + IIF(aisoms.s_pred - aisoms.sum_flk - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b>0, aisoms.s_pred - aisoms.sum_flk - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b, 0)
 m.sum13 = m.sum13 + IIF(aisoms.s_pred - aisoms.sum_flk - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b<0, -(aisoms.s_pred - aisoms.sum_flk - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b), 0)

 m.nincokr = m.nincokr + 1

 IF cokr != m.ccokr
  m.fstring = m.fstring + 1
  oExcel.Range(oExcel.Cells(m.fstring,2), oExcel.Cells(m.fstring,3)).Merge 
  oExcel.Sheets(1).Cells(m.fstring,02).Value2 = '����� �� ������:'
  WITH oExcel.Sheets(1)
   .Cells(m.fstring,05).Value2 = m.sum04okr
   .Cells(m.fstring,06).Value2 = m.sum05okr
   .Cells(m.fstring,07).Value2 = m.sum06okr
   .Cells(m.fstring,08).Value2 = m.sum07okr
   .Cells(m.fstring,09).Value2 = m.sum08okr
   .Cells(m.fstring,10).Value2 = m.sum09okr
   .Cells(m.fstring,11).Value2 = m.sum10okr
   .Cells(m.fstring,12).Value2 = m.sum11okr
   .Cells(m.fstring,13).Value2 = m.sum12okr
   .Cells(m.fstring,14).Value2 = m.sum13okr
  ENDWITH 

  m.fstring = m.fstring + 1
  m.ccokr = cokr
  m.cokrname = IIF(SEEK(m.ccokr, 'admokr'), admokr.name_okr, '')
  oExcel.Sheets(1).cells(m.fstring,2).Value2 = m.cokrname
  oRange = oExcel.Range(oExcel.Cells(m.fstring,2), oExcel.Cells(m.fstring,14))
  oRange.Merge
  oRange.HorizontalAlignment = -4108
  oRange.Interior.ColorIndex = 40
  m.nincokr = 1
  m.cokrsum = 0

  m.sum04okr = 0
  m.sum05okr = 0
  m.sum06okr = 0
  m.sum07okr = 0
  m.sum08okr = 0
  m.sum09okr = 0
  m.sum10okr = 0
  m.sum11okr = 0
  m.sum12okr = 0
  m.sum13okr = 0
 ENDIF 
ENDSCAN 

m.fstring = m.fstring + 1
WITH oExcel.Sheets(1)
  oExcel.Range(oExcel.Cells(m.fstring,2), oExcel.Cells(m.fstring,3)).Merge 
 .Cells(m.fstring,02).Value2 = '����� �� ���� �������:'
 .Cells(m.fstring,05).Value2 = m.sum04
 .Cells(m.fstring,06).Value2 = m.sum05
 .Cells(m.fstring,07).Value2 = m.sum06
 .Cells(m.fstring,08).Value2 = m.sum07
 .Cells(m.fstring,09).Value2 = m.sum08
 .Cells(m.fstring,10).Value2 = m.sum09
 .Cells(m.fstring,11).Value2 = m.sum10
 .Cells(m.fstring,12).Value2 = m.sum11
 .Cells(m.fstring,13).Value2 = m.sum12
 .Cells(m.fstring,14).Value2 = m.sum13
ENDWITH 

IF fso.FileExists(pDir+'\ActSF.xls')
 fso.DeleteFile(pDir+'\ActSF.xls')
ENDIF 

oBook.SaveAs(BookName,18)

oExcel.Visible = .T.

SET RELATION OFF INTO aisoms
USE IN aisoms
USE IN sprlpu
USE IN admokr
IF USED('lpudogs')
 USE IN lpudogs
ENDIF 

WAIT CLEAR 

RETURN 
