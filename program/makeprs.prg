FUNCTION MakePrs(para1)
 PRIVATE mcod,mmy,lpu_id,IsLpuTpn
 m.mcod   = para1
 m.mmy    = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
 m.lpu_id = lpuid
 m.IsLpuTpn = IIF(SEEK(m.lpu_id, 'lputpn'), .t., .f.)
 m.IsPilot = IIF(SEEK(m.mcod, 'pilot'), .t., .f.)
 m.IsPilotS= IIF(SEEK(m.mcod, 'pilots'), .t., .f.)

 m.sum_st1 = 0  && ����� � ������, ���������� ���������� ��� �� ��������������
 m.sum_st1 = s_pred - sum_flk

 m.nIsDoubles = 0

 lcPath = pBase+'\'+gcperiod+'\'+mcod
 IF !fso.FolderExists(lcPath)
  MESSAGEBOX(CHR(13)+CHR(10)+'����������� ���������� ���!'+CHR(13)+CHR(10),0+64,mcod)
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\TarifN', 'tarif', 'shar', 'cod')>0
  IF USED('TARIF')
   USE IN tarif
  ENDIF 
  RETURN 
 ENDIF 

 =MakeYFilesOne(lcPath)
 
 USE IN tarif
 SELECT aisoms

 
 MESSAGEBOX("��������� ���������!",0+64,"")
 
RETURN 