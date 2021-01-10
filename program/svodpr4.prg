PROCEDURE SvodPr4
 IF MESSAGEBOX(CHR(13)+CHR(10)+'�� ������ ������������ ���� ���������� 4?'+CHR(13)+CHR(10)+;
 '(������� 1)'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\pr4.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'����������� ���� PR4.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\pr4', 'pr4', 'shar')>0
  IF USED('pr4')
   USE IN pr4
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
  IF USED('pr4')
   USE IN pr4
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod') > 0
  IF USED('pr4')
   USE IN pr4
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\lputpn', "lputpn", "shar", "lpu_id") > 0
  IF USED('lputpn')
   USE IN lputpn
  ENDIF 
  IF USED('pr4')
   USE IN pr4
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pcommon+'\lpudogs', 'lpudogs', 'shar', 'lpu_id') > 0
  IF USED('lputpn')
   USE IN lputpn
  ENDIF 
  IF USED('lpudogs')
   USE IN lpudogs
  ENDIF 
  IF USED('pr4')
   USE IN pr4
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\FormMAG02', 'mag', 'shar', 'lpuid') > 0
  IF USED('mag')
   USE IN mag
  ENDIF 
  USE IN lputpn
  USE IN pr4
  USE IN sprlpu
  USE IN aisoms
  USE IN lpudogs
  RETURN 
 ENDIF 
 
 CREATE CURSOR svodpr4 (mcod c(7), lpuid n(4), mek n(11,3))

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
 oExcel.ReferenceStyle= -4150  && xlR1C1
 oExcel.SheetsInNewWorkbook = 1
 oBook = oExcel.WorkBooks.Add
 oSheet = oBook.WorkSheets(1)
 oSheet.Select
 oSheet.Name = '�������'
 oexcel.Cells.Font.Name='Calibri'
 oexcel.ActiveSheet.PageSetup.Orientation=2
 oExcel.Cells.NumberFormat = '@'

 oRange = oExcel.Range(oExcel.Cells(1,1), oExcel.Cells(1,18))
 oRange.Merge
 oExcel.Cells(1,1).Value='������� ��� �� ������ �������� �� ���������� ��������������'
 oExcel.Cells(1,1).HorizontalAlignment = -4108
 oExcel.Cells(1,1).Font.Size = 12
 oExcel.Cells(1,1).Font.Bold = .F.
 oExcel.Cells(1,1).Font.Italic = .T.
 oExcel.Rows(1).RowHeight = 30
 oExcel.Rows(1).VerticalAlignment = -4108
 
 oExcel.Columns(1).ColumnWidth  = 3
 oExcel.Columns(2).ColumnWidth  = 21
 oExcel.Columns(3).ColumnWidth  = 8
 oExcel.Columns(4).ColumnWidth  = 12
 oExcel.Columns(5).ColumnWidth  = 8
 oExcel.Columns(6).ColumnWidth  = 13

 oExcel.Columns(7).ColumnWidth  = 13
 oExcel.Columns(8).ColumnWidth  = 13
 oExcel.Columns(9).ColumnWidth  = 13
 oExcel.Columns(10).ColumnWidth = 13
 oExcel.Columns(11).ColumnWidth = 13
 oExcel.Columns(12).ColumnWidth = 13
 oExcel.Columns(13).ColumnWidth = 13
 oExcel.Columns(14).ColumnWidth = 13
 oExcel.Columns(15).ColumnWidth = 13
 oExcel.Columns(16).ColumnWidth = 13
 oExcel.Columns(17).ColumnWidth = 13
 oExcel.Columns(18).ColumnWidth = 13
 oExcel.Columns(19).ColumnWidth = 13
 oExcel.Columns(20).ColumnWidth = 13
 oExcel.Columns(21).ColumnWidth = 13
 
 oExcel.Cells(2,1).Value  = '� �\�'
 oExcel.Cells(2,2).Value  = '������������ ��� (������������ ����) � ������� �� �������'
 oExcel.Cells(2,3).Value  = '�����'
 oExcel.Cells(2,4).Value  = '��� �� (mcod)'
 *oExcel.Cells(2,5).Value  = '����� �������� � ��'
 oExcel.Cells(2,5).Value  = '��� �� (lpuid)'
 oExcel.Cells(2,6).Value  = '���-�� �������������� ���������'

 oExcel.Cells(2,7).Value  = '��������� ����� ���������� ��������������'
 oExcel.Cells(2,8).Value  = '���'
 oExcel.Cells(2,9).Value  = '����� �������� ������'
 oExcel.Cells(2,10).Value = '����� �������, ���������� ���������� �� �������������� �� ;
  ��������� ������������� � ������� �� �� ��������, ������������� � ������� ��, ��;
  ���������� ������������ ����������� ������ � ������ ��'
 oExcel.Cells(2,11).Value = '����� �������, ���������� �� ���������, ������������� ;
  � ������ ��, �� � �������� ������� ����������� � ������ ��'
 oExcel.Cells(2,12).Value  = '����� �������������� �������,���������� �� �� ����������� ������ ;
  ����������������� ����������� � ������� �� ��������� ����������� ������ � ������� ;
  �����������'
 oExcel.Cells(2,13).Value  = '����� ������� �� ����������� ������,��������� � ������ �� ;
  ���������,������������� � ��,�� ����������� � ��������� ��������������'
 oExcel.Cells(2,14).Value  = '����� ������� �� ����������� ������, ��������� � ������ �� ;
  ���������, �� ������������� � ��������� ������������'
 oExcel.Cells(2,15).Value = '����� ������� �� ����������� ���������� ���'
 oExcel.Cells(2,16).Value = '����� ������� �� ����������� ���������� ����'
 oExcel.Cells(2,17).Value = '��������� ��������'
 oExcel.Cells(2,18).Value = '���� �� ������ �������'
 oExcel.Cells(2,19).Value = '����� � ������'
 oExcel.Cells(2,20).Value = '���� �� ����� �������'
 oExcel.Cells(2,21).Value = '������� �� ��'

 oExcel.Rows(2).RowHeight = 130
 oExcel.Rows(2).HorizontalAlignment = 1
 oExcel.Rows(2).VerticalAlignment = -4160
 oExcel.Rows(2).WrapText = .t.

 oExcel.Cells(3,1).Value  = '1'
 oExcel.Cells(3,2).Value  = '2'
 oExcel.Cells(3,3).Value  = '3'
 oExcel.Cells(3,4).Value  = '4'
 oExcel.Cells(3,5).Value  = '5'
 oExcel.Cells(3,6).Value  = '6'
 oExcel.Cells(3,7).Value  = '7'
 oExcel.Cells(3,8).Value  = '8'
 oExcel.Cells(3,9).Value  = '9'
 oExcel.Cells(3,10).Value = '10'
 oExcel.Cells(3,11).Value = '11'
 oExcel.Cells(3,12).Value = '12'
 oExcel.Cells(3,13).Value = '13'
 oExcel.Cells(3,14).Value = '14'
 oExcel.Cells(3,15).Value = '15'
 oExcel.Cells(3,16).Value = '16'
 oExcel.Cells(3,17).Value = '17'
 oExcel.Cells(3,18).Value = '18'
 oExcel.Cells(3,19).Value = '19'
 oExcel.Cells(3,20).Value = '20'
 oExcel.Cells(3,21).Value = '21'

 nRow  = 4
 nnRow = 1

 m.col6  = 0
 m.col7  = 0
 m.col8  = 0
 m.col9  = 0
 m.col10 = 0
 m.col11 = 0
 m.col12 = 0
 m.col13 = 0
 m.col14 = 0
 m.col15 = 0
 m.col16 = 0
 m.col17 = 0
 m.col18 = 0
 m.col19 = 0
 m.col20 = 0
 m.col21 = 0

 SELECT pr4
 SET RELATION TO mcod INTO aisoms 
 IF USED('mag')
  SET RELATION TO lpuid INTO mag ADDITIVE 
 ENDIF 
 SCAN 
  SCATTER MEMVAR 
  m.lpuname = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
  m.cokr = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.cokr, '')
  m.lpudog = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.dogs), '')
  m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)

  *IF m.IsLpuTpn
  * m.s_avans = aisoms.s_pr_avans
  *ELSE 
  * m.s_avans = aisoms.s_avans
  *ENDIF 
  *m.s_avans = aisoms.s_pr_avans
  
  IF aisoms.s_avans=0 AND aisoms.s_pr_avans=0
   m.s_avans = 0
  ENDIF 
  IF aisoms.s_avans>0 AND aisoms.s_pr_avans=0
   m.s_avans = aisoms.s_avans
  ENDIF 
  IF aisoms.s_avans>0 AND aisoms.s_pr_avans>0
   m.s_avans = aisoms.s_pr_avans
  ENDIF 

  m.str01 = aisoms.finval
  *m.str02 = aisoms.s_pr_avans 
  m.str02 = m.s_avans 
  m.str33 = s_own - m.str02
  m.str31 = s_others
  m.str32 = s_guests
  m.str04 = s_kompl + s_dst
  m.str06 = s_npilot
  m.str07 = s_empty
  m.str03 = m.str02 - m.str31 + m.str32 + m.str33
  m.str05 = m.str01 - IIF(m.str03<0, -1*m.str03, m.str03)
  *m.str08 = aisoms.sum_flk
  *m.str08 = aisoms.pf_flk
  m.str08 = IIF(USED('mag'), mag.col16, 0)
  m.str09 = IIF(m.str05>=0, m.str05, 0)+m.str04+m.str06+m.str07
  m.defs  = aisoms.e_mee+aisoms.e_ekmp
  m.koplate = m.str01 - m.str02 - m.str31 + m.str32 + m.str04 + m.str06 + m.str07 - m.defs

  oExcel.Cells(nRow,1).Value  = nnRow
  oExcel.Cells(nRow,2).Value  = m.lpuname
  oExcel.Cells(nRow,3).Value  = m.cokr
  oExcel.Cells(nRow,4).Value  = m.mcod
  *oExcel.Cells(nRow,5).Value  = m.lpudog
  oExcel.Cells(nRow,5).Value  = STR(m.lpuid,4)

  oExcel.Cells(nRow,6).Value  = aisoms.pazval
  oExcel.Cells(nRow,7).Value  = aisoms.finval
  oExcel.Cells(nRow,8).Value  = m.str08
  *oExcel.Cells(nRow,9).Value  = aisoms.s_pr_avans
  oExcel.Cells(nRow,9).Value  = m.s_avans
  oExcel.Cells(nRow,10).Value = m.str31
  oExcel.Cells(nRow,11).Value = m.str32
  oExcel.Cells(nRow,12).Value = m.str04
  oExcel.Cells(nRow,13).Value = m.str06
  oExcel.Cells(nRow,14).Value = m.str07 
  oExcel.Cells(nRow,15).Value = aisoms.e_mee
  oExcel.Cells(nRow,16).Value = aisoms.e_ekmp
  oExcel.Cells(nRow,17).Value = s_bad
  oExcel.Cells(nRow,18).Value = aisoms.dolg_b
  oExcel.Cells(nRow,19).Value = IIF(m.koplate>=0, m.koplate, 0)
  oExcel.Cells(nRow,20).Value = IIF(m.koplate<0, -1*m.koplate, 0) && ���� �� ����� �������
  oExcel.Cells(nRow,21).Value = IIF(FIELD('s_lek', 'aisoms')='S_LEK', aisoms.s_lek, 0)

  m.col6  = m.col6 + aisoms.pazval
  m.col7  = m.col7 + aisoms.finval
  m.col8  = m.col8 + m.str08
  *m.col9  = m.col9 + aisoms.s_pr_avans
  m.col9  = m.col9 + m.s_avans
  m.col10 = m.col10 + m.str31
  m.col11 = m.col11 + m.str32
  m.col12 = m.col12 + m.str04
  m.col13 = m.col13 + m.str06
  m.col14 = m.col14 + m.str07
  m.col15 = m.col15 + aisoms.e_mee
  m.col16 = m.col16 + aisoms.e_ekmp
  m.col17 = m.col17 + s_bad
  m.col18 = m.col18 + aisoms.dolg_b
  m.col19 = m.col19 + IIF(m.koplate>=0, m.koplate, 0)
  m.col20 = m.col20 + IIF(m.koplate<0, -1*m.koplate, 0)
  m.col21 = m.col21 + IIF(FIELD('s_lek', 'aisoms')='S_LEK', aisoms.s_lek, 0)

  nRow  = nRow + 1
  nnRow = nnRow + 1
  
  INSERT INTO svodpr4 (mcod,lpuid,mek) VALUES (m.mcod,m.lpuid,m.str08)
  
 ENDSCAN 
 SET RELATION OFF INTO aisoms
 IF USED('mag')
  SET RELATION OFF INTO mag 
 ENDIF 
 
 IF USED('pr4')
  USE IN pr4
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 IF USED('lputpn')
  USE IN lputpn
 ENDIF 
 IF USED('mag')
  USE IN mag
 ENDIF 
 
 SELECT svodpr4
 COPY TO &pBase\&gcPeriod\svodpr4
 USE 
 
 WAIT CLEAR 
 
 oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,5))
 oRange.Merge
 oExcel.Cells(nRow,1).Value='�����:'

 oExcel.Cells(nRow,6).Value  = m.col6
 oExcel.Cells(nRow,7).Value  = m.col7
 oExcel.Cells(nRow,8).Value  = m.col8
 oExcel.Cells(nRow,9).Value  = m.col9
 oExcel.Cells(nRow,10).Value = m.col10
 oExcel.Cells(nRow,11).Value = m.col11
 oExcel.Cells(nRow,12).Value = m.col12
 oExcel.Cells(nRow,13).Value = m.col13
 oExcel.Cells(nRow,14).Value = m.col14
 oExcel.Cells(nRow,15).Value = m.col15
 oExcel.Cells(nRow,16).Value = m.col16
 oExcel.Cells(nRow,17).Value = m.col17
 oExcel.Cells(nRow,19).Value = m.col19
 oExcel.Cells(nRow,20).Value = m.col20
 oExcel.Cells(nRow,21).Value = m.col21

 BookName = 'svpr4'+m.qcod+PADL(DAY(DATE()),2,'0')+PADL(MONTH(DATE()),2,'0')
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+BookName+'.xls')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+BookName+'.xls')
 ENDIF 

 oBook.SaveAs(pbase+'\'+m.gcperiod+'\'+BookName+'.xls',18)
 oExcel.Visible = .T.
 
RETURN 