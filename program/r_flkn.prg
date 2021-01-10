PROCEDURE r_flkn
 * ������������ ������������ �� ����� people � ������� �� ��������� ���� r_flk!

 *-- ���������� ������� � ������� ������
 lcAdmitSymb = ''  
 *-- ���� �� ���� ������� ������ �� � �� �
 FOR li=192 To 255  
 	lcAdmitSymb = lcAdmitSymb + Chr(li)  
 ENDFOR 
 *-- ������� ���� 32 ������, 45 ����, 168 �, 184 �, 39 ', ` 96 ', 46 .
 lcAdmitSymb = lcAdmitSymb + CHR(32)+CHR(45)+CHR(168)+CHR(184)+CHR(39)+CHR(96)+CHR(46)
 *? CHRTRAN(m.Fam, CHRTRAN(m.Fam, lcAdmitSymb,''), '') && ���� ����� m.Fam, �� ��� ��!
 
 IF M.ERA == .T. && �������� ER
  IF !EMPTY(sv)  
   m.IsGood = IIF(SEEK(sv, 'osoerz') AND osoerz.kl == 'y', .T., .F.)
   IF IsVS(sn_pol) AND LEFT(sn_pol,2)=m.qcod
    IF USED('kms')
     m.vvs = INT(VAL(SUBSTR(ALLTRIM(sn_pol),7)))
     IF SEEK(m.vvs, 'kms')
      m.IsGood = .t.
     ENDIF 
    ENDIF 
   ENDIF 
   IF !IsGood
    m.recid = recid
    rval = InsError('S', 'PKA', m.recid, '',;
    	'������ ����� ����������� �� ����������� ������ ERA')
    *m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
   ENDIF 
  ENDIF 
 ENDIF 

 IF M.ECA == .T. && �������� EC
  IF !EMPTY(sv)
   m.IsGood = IIF(IIF(m.qcod<>'S7', qq = m.qcod, qq=m.qcod OR qq='R2'), .T., .F.)
   IF IsVS(sn_pol) AND LEFT(sn_pol,2)=m.qcod
    IF USED('kms')
     m.vvs = INT(VAL(SUBSTR(ALLTRIM(sn_pol),7)))
     IF SEEK(m.vvs, 'kms')
      m.IsGood = .t.
     ENDIF 
    ENDIF 
   ENDIF 
   IF !IsGood
    m.recid = recid
    =InsError('R', 'ECA', m.recid, '',;
    	'������ ��������� �������������� ��������� ���')
   ENDIF 
  ENDIF 
 ENDIF 
   
 IF M.E1A == .T.  && �������� E1
  m.d_type = d_type
  m.recid  = recid
  m.fam    = fam
  m.im     = im
  m.ot     = ot

  DO CASE 
   CASE EMPTY(m.d_type)
    =InsError('R', 'E1A', m.recid, '',;
    	'������ ���� d_type')
   CASE !SEEK(m.d_type, 'osoree')
    =InsError('R', 'E1A', m.recid, '',;
    	'������������ �������� ���� d_type')

   OTHERWISE 
    IF INLIST(m.d_type,'4','t') AND !(EMPTY(m.Fam) AND !EMPTY(m.Ot))
     =InsError('R', 'E1A', m.recid,'',;
   	 	'd_type='+m.d_type+' ��� ������� �������')
    ENDIF 
    IF INLIST(m.d_type,'6','m') AND !(EMPTY(m.Fam) AND EMPTY(m.Ot))
     =InsError('R', 'E1A', m.recid,'',;
   	 	'd_type='+m.d_type+' ��� ������� �������')
    ENDIF 
    IF m.d_type='U' AND LEN(ALLTRIM(m.fam))<>1 
     =InsError('R', 'E1A', m.recid,'',;
   	 	'd_type='+m.d_type+' ��� �������, �� ��������� �� ������ �������')
    ENDIF 
    IF INLIST(m.d_type,'5','k') AND !(EMPTY(m.Im) OR EMPTY(m.Ot))
     =InsError('R', 'E1A', m.recid,'',;
   	 	'd_type='+m.d_type+' ��� ������� ����� �/��� ��������')
    ENDIF 

  ENDCASE 
 ENDIF 
   
 IF M.E2A == .T. && �������� E2
  m.recid = recid
  IF !INLIST(tipp, '�','�','�','�','�')
   =InsError('R', 'E2A', m.recid, '',;
   	"������������ ���� tip_p ('�','�','�','�','�')")
  ELSE  
   DO CASE 
    CASE tipp='�'
     IF !IsKms(people.sn_pol)
     =InsError('R', 'E2A', m.recid, '',;
   	 	'�������� ������ ������ ���� "�"')
   	 ENDIF 

    CASE tipp='�'
     IF !IsVS(people.sn_pol) AND !IsVSN(people.sn_pol)&& ��� � �������� - 3-� ������� ��� ���+��+������+ �����
      =InsError('R', 'E2A', m.recid, '',;
   	 	'�������� ������ ������ ���� "�": ��� ���+��� ��+����� (9 ����)')
   	 ENDIF 
   	 	
   	OTHERWISE 
   	 IF !IsENP(sn_pol)
      =InsError('R', 'E2A', m.recid, '',;
   	 	'�������� ������ ������ ���� "'+tipp+'" (16 �������� �������� �� �������� ����� �����)')
   	 ENDIF 

   ENDCASE 

  ENDIF 

  *IF (!IsKms(sn_pol) AND !IsVS(sn_pol) AND !IsVSN(sn_pol) AND !IsENP(sn_pol))
  * m.recid = recid
  * =InsError('R', 'E2A', m.recid, '',;
  * 	'������������ �������� ���')
  *ENDIF 
 ENDIF 

 *IF  M.E4A == .T. AND IIF(!INLIST(m.qcod,'R2','S7'), .T., .F.) && �������� E4
 * IF ((INLIST(RIGHT(PADL(ALLTRIM(fam),25),2),'��','��','��') AND INLIST(RIGHT(PADL(ALLTRIM(ot),20),2),'��','��') AND w!=2) OR ;
 *    (INLIST(RIGHT(PADL(ALLTRIM(fam),25),2),'��','��','��')  AND INLIST(RIGHT(PADL(ALLTRIM(ot),20),2),'��','��') AND w!=1))
 *  m.recid = recid
 *  =InsError('R', 'E4A', m.recid,'',;
 *  	'������������� ��� ���� (��������, ��� ������ �������, � ���, ��������, �������, ��� ��������)')
 * ENDIF 
 *ENDIF 
   
 IF  M.E4A == .T. && �������� E4
  m.recid = recid
  m.Fam    = Fam
  m.Ot     = Ot
  m.d_type = d_type

  IF (EMPTY(m.Fam) AND !EMPTY(m.ot)) AND !INLIST(m.d_type,'4','t')
   =InsError('R', 'E4A', m.recid,'',;
   	'���������� ������� ��� ������� �������� �� ������������ d_type=4,t')
  ENDIF 
  IF (EMPTY(m.Fam) AND EMPTY(m.ot)) AND !INLIST(m.d_type,'6','m')
   =InsError('R', 'E4A', m.recid,'',;
   	'���������� ������� � �������� �� ������������ d_type=6,m')
  ENDIF 
  IF LEN(ALLTRIM(m.fam))=1 AND !INLIST(m.d_type,'U','9','2')
   =InsError('R', 'E4A', m.recid,'',;
   	'�������, ��������� �� ����� �����, �� ������������ d_type=U,9')
  ENDIF 
  IF CHRTRAN(m.Fam, CHRTRAN(m.Fam, lcAdmitSymb+CHR(95),''), '') <> m.Fam
   =InsError('R', 'E4A', m.recid,'',;
   	'������������ ������ � �������!')
  ENDIF 
 ENDIF 
 
 IF M.E5A
  m.recid = recid
  m.Im    = Im
  m.Ot    = Ot
  m.d_type = d_type
  IF (EMPTY(m.Im) AND EMPTY(m.Ot)) AND !INLIST(m.d_type,'5','k')
   =InsError('R', 'E5A', m.recid,'',;
   	'���������� ����� � �������� �� ������������ d_type=5,k')
  ENDIF 
  IF CHRTRAN(m.Im, CHRTRAN(m.Im, lcAdmitSymb,''), '') <> m.Im
   =InsError('R', 'E5A', m.recid,'',;
   	'������������ ������ � �����!')
  ENDIF 
 ENDIF 

 IF  M.E6A == .T. && �������� E6
  m.recid = recid
  m.Fam   = Fam
  m.Im    = Im
  m.Ot    = Ot
  m.d_type = d_type
  
  IF (EMPTY(m.Ot) AND !EMPTY(m.Im)) AND !INLIST(m.d_type,'9','2','f','6','m','U')
   =InsError('R', 'E6A', m.recid,'',;
   	'���������� �������� ��� d_type!=9,2,f,6,m,U')
  ENDIF 
  IF (EMPTY(m.Ot) AND EMPTY(m.Im)) AND !INLIST(m.d_type,'5','k')
   =InsError('R', 'E6A', m.recid,'',;
   	'���������� ����� � �������� ��� d_type!=5,k')
  ENDIF 
  IF (EMPTY(m.Ot) AND EMPTY(m.Fam)) AND !INLIST(m.d_type,'6','m')
   =InsError('R', 'E6A', m.recid,'',;
   	'���������� ����� � ������� ��� d_type!=6,m')
  ENDIF 
  IF CHRTRAN(m.Ot, CHRTRAN(m.Ot, lcAdmitSymb,''), '') <> m.Ot
   =InsError('R', 'E6A', m.recid,'',;
   	'������������ ������ � ��������!')
  ENDIF 
  
  *IF (EMPTY(ot) AND !INLIST(d_type,'2','f','9','U')) OR ;
  *	(!EMPTY(ot) AND INLIST(d_type,'2'))
  * m.recid = recid
  * =InsError('R', 'E6A', m.recid,'',;
  * 	'���������� �������� ��� d_type!=2,f')
  *ENDIF 
 ENDIF 
   
 IF M.E7A == .T. && �������� E7
  IF (!INLIST(w,1,2) OR (IsKms(sn_pol) AND SUBSTR(sn_pol,5,2)!='77' AND (w != IIF(VAL(SUBSTR(sn_pol,12,2))>50, 1, 2))))
   m.recid = recid
   =InsError('R', 'E7A', m.recid)
  ENDIF 
 ENDIF 

 IF M.E7A == .T.
  m.sn_pol = sn_pol                && �������� E7
  Dtt = CTOD(IIF(VAL(SUBSTR(m.sn_pol,12,2))>50, ;
       PADL(INT(VAL(SUBSTR(m.sn_pol,12,2))-50),2,'0'), ;
       SUBSTR(m.sn_pol,12,2))+'.'+IIF(VAL(SUBSTR(m.sn_pol,14,2))>40, ;
       PADL(INT(VAL(SUBSTR(m.sn_pol,14,2))-40),2,'0')+'.20', ;
       SUBSTR(m.sn_pol,14,2)+'.19')+SUBSTR(m.sn_pol,16,2))
  IF (IsKms(m.sn_pol) AND !INLIST(SUBSTR(m.sn_pol,5,2),'50','51') AND (dr != Dtt))
   m.recid = recid
   =InsError('R', 'E7A', m.recid)
  ENDIF 
 ENDIF 

 IF M.E8A == .T.
  m.sn_pol = people.sn_pol                && �������� E8
  IF (people.dr=={} OR (dat1-IIF(!EMPTY(people.dr), people.dr, {01.01.1850}))/365.25>120 OR ;
   IIF(!EMPTY(people.dr), people.dr, {01.01.1850}) > m.dat2)
   m.recid = people.recid
   =InsError('R', 'E8A', m.recid)
  ENDIF 
 ENDIF 
  
RETURN 