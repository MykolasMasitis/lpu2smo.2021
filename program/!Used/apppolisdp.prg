PROCEDURE AppPolisDP
 IF MESSAGEBOX(''+CHR(13)+CHR(10)+'�� ������ �������� '+;
  '���������� POLIC_DP?',4+32,'')==7
  RETURN 
 ENDIF 
 
 oal = SYS(5)+SYS(2003)
 SET DEFAULT TO (pbase+'\'+gcperiod+'\'+'nsi')
 dpnew = GETFILE('dbf','','',0,'������� �� ����!')
 SET DEFAULT TO (oal)
 
 IF EMPTY(dpnew)
  MESSAGEBOX('�� ������ �� �������!',0+64,'')
  RETURN 
 ENDIF 
 
 tnresult = 0
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\polic_dp', 'dp_svod', 'shared', 'sn_pol')
 tnresult = tnresult + OpenFile(dpnew, 'dp_new', 'shar')
 
 IF tnresult>0
  IF USED('dp_svod')
   USE IN dp_svod
  ENDIF 
  IF USED('dp_new')
   USE IN dp_new
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT dp_new
 IF VARTYPE(sn_pol)!='C'
  USE
  USE IN dp_svod
  MESSAGEBOX('� ����� '+dpnew+CHR(13)+CHR(10)+;
   '������������ ����������� ���� SN_POL!'+CHR(13)+CHR(10)+;
   '����������� ������ ����������!',0+16, '')
  RETURN 
 ENDIF 
 IF VARTYPE(D_U)!='D'
  USE 
  USE IN dp_svod
  MESSAGEBOX('� ����� '+dpnew+CHR(13)+CHR(10)+;
   '������������ ����������� ���� D_U!'+CHR(13)+CHR(10)+;
   '����������� ������ ����������!',0+16, '')
  RETURN 
 ENDIF 
 
 SELECT dp_new
 m.totrecs = RECCOUNT()
 m.addrecs = 0
 PRIVATE lpu_id, sn_pol, qq, d_u, tms, year
 SCAN 
  SCATTER MEMVAR 
*  m.sn_pol = sn_pol
*  m.d_u = d_u
  IF !SEEK(m.sn_pol, 'dp_svod')
*   INSERT INTO dp_svod (sn_pol, d_u) VALUES (m.sn_pol, m.d_u)
   INSERT INTO dp_svod FROM MEMVAR 
   m.addrecs = m.addrecs + 1
  ELSE 
  ENDIF 
 ENDSCAN 
 USE 
 USE IN dp_svod
 
 MESSAGEBOX('��������� '+ALLTRIM(STR(m.addrecs))+ ' ������� � ���� POLIC_DP'+CHR(13)+CHR(10)+;
  '�� '+ALLTRIM(STR(m.totrecs))+ ' � ��� ��������������!',0+64, '')

RETURN 