PROCEDURE r_flk
 * сканирование производится по файл talon!
   IF M.ERA == .T. && Алгоритм ER
    IF !EMPTY(people.sv)  
     m.IsGood = IIF(SEEK(people.sv, 'osoerz') AND osoerz.kl == 'y', .T., .F.)
     IF IsVS(people.sn_pol) AND LEFT(people.sn_pol,2)=m.qcod
      IF USED('kms')
       m.vvs = INT(VAL(SUBSTR(ALLTRIM(people.sn_pol),7)))
       IF SEEK(m.vvs, 'kms')
        m.IsGood = .t.
       ENDIF 
      ENDIF 
     ENDIF 
     IF IsGood == .f.
      m.polis = sn_pol
      m.recsproc = 0 
      DO WHILE sn_pol == m.polis
       m.recid = recid
       rval = InsError('S', 'PKA', m.recid, '',;
       	'Запись счета забракована по регистровой ошибке ERA')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       SKIP +1 
       m.recsproc = m.recsproc + 1
      ENDDO 
      SKIP -1*(m.recsproc)
      m.recid = people.recid
      =InsError('R', 'ERA', m.recid, '', ;
      	'Отрицательный вектор сверки')
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.ECA == .T. && Алгоритм EC
    IF !EMPTY(people.sv)
     m.IsGood = IIF(people.qq = m.qcod, .T., .F.)
     IF IsVS(people.sn_pol) AND LEFT(people.sn_pol,2)=m.qcod
      IF USED('kms')
       m.vvs = INT(VAL(SUBSTR(ALLTRIM(people.sn_pol),7)))
       IF SEEK(m.vvs, 'kms')
        m.IsGood = .t.
       ENDIF 
      ENDIF 
     ENDIF 
     IF IsGood == .f.                 
      m.polis = sn_pol
      m.recsproc = 0 
      DO WHILE sn_pol == m.polis
       m.recid = recid
       rval =InsError('S', 'PKA', m.recid, '',;
       	'Услуга забракована по регистровой ошибке ECA')
*       InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       SKIP +1 
       m.recsproc = m.recsproc + 1
      ENDDO 
      SKIP -1*(m.recsproc)
      m.recid = people.recid
      =InsError('R', 'ECA', m.recid, '',;
      	'Ошибка страховой принадлежности документа ОМС')
*      InsErrorSV(m.mcod, 'R', 'ECA', m.recid)
     ENDIF 
    ENDIF 
   ENDIF 
   
   IF M.E1A == .T.  && Алгоритм E1
    DO CASE 
     CASE EMPTY(people.d_type)
      m.polis = sn_pol
      m.recsproc = 0 
      DO WHILE sn_pol == m.polis
       m.recid = recid
       rval =InsError('S', 'PKA', m.recid, '', ;
       	'Запись счета забракована по регистровой ошибке E1A (пустое поле d_type)')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       SKIP +1 
       m.recsproc = m.recsproc + 1
      ENDDO 
      SKIP -1*(m.recsproc)
      m.recid = people.recid
      =InsError('R', 'E1A', m.recid, '', ;
      	'Поле d_type не заполнено')

     CASE !SEEK(people.d_type, 'osoree')
      m.polis = sn_pol
      m.recsproc = 0 
      DO WHILE sn_pol == m.polis
       m.recid = recid
       rval =InsError('S', 'PKA', m.recid, '', ;
       	'Запись счета забракована по регистровой ошибке E1A (значение поля d_type отсутствует в справочнике osoree)')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       SKIP +1 
       m.recsproc = m.recsproc + 1
      ENDDO 
      SKIP -1*(m.recsproc)
      m.recid = people.recid
      =InsError('R', 'E1A', m.recid, '2')
      
      * 3. Значение поля не соответствует указанной категории пациента.
      * Как проверять?!

     OTHERWISE 
    ENDCASE 
   ENDIF 
   
   IF M.E2A == .T. && Алгоритм E2
    IF !INLIST(tip_p, 'С','П','В','К','Э')
    ELSE  
     DO CASE 
      CASE tip_p='С' AND !IsKms(people.sn_pol)
     ENDCASE 
    ENDIF 

   
    IF (!IsKms(people.sn_pol) AND !IsVS(people.sn_pol) AND !IsVSN(people.sn_pol) AND !IsENP(people.sn_pol))
     m.polis = sn_pol
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E2A', m.recid)
*     InsErrorSV(m.mcod,'R', 'E2A', m.recid)
    ENDIF 
   ENDIF 

   IF  M.E4A == .T. AND IIF(!INLIST(m.qcod,'R2','S7'), .T., .F.) && Алгоритм E4
    IF ((INLIST(RIGHT(PADL(ALLTRIM(People.fam),25),2),'ва','на','ая') AND INLIST(RIGHT(PADL(ALLTRIM(People.ot),20),2),'на','зы') AND People.w!=2) OR ;
       (INLIST(RIGHT(PADL(ALLTRIM(People.fam),25),2),'ов','ев','ин')  AND INLIST(RIGHT(PADL(ALLTRIM(People.ot),20),2),'ич','лы') AND People.w!=1))
     m.polis = sn_pol 
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid,'',;
      	'Услуга забракована по регистровой ошибке E4A')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E4A', m.recid,'',;
     	'Несоответсвие ФИО полу (вероятно, пол указан мужской, а ФИО, очевидно, женская, или наоборот)')
    ENDIF 
   ENDIF 
   
   IF  M.E4A == .T. && Алгоритм E4
    DO CASE 
     CASE LEN(ALLTRIM(people.fam))=1 AND people.d_type<>'U'
      m.polis = sn_pol 
      m.recsproc = 0 
      DO WHILE sn_pol == m.polis
       m.recid = recid
       rval =InsError('S', 'PKA', m.recid,'',;
      	'Услуга забракована по регистровой ошибке E4A')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       SKIP +1 
       m.recsproc = m.recsproc + 1
      ENDDO 
      SKIP -1*(m.recsproc)
      m.recid = people.recid
      =InsError('R', 'E4A', m.recid,'',;
     	'Фамилия, состоящая из одной буквы, не подтверждена d_type="U"')
     OTHERWISE 
    ENDCASE 
   ENDIF 

   IF  M.E6A == .T. && Алгоритм E4
    IF (EMPTY(people.ot) AND !INLIST(people.d_type,'2','f','9','U')) OR ;
    	(!EMPTY(people.ot) AND INLIST(people.d_type,'2'))
     m.polis = sn_pol 
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid,'',;
      	'Услуга забракована по регистровой ошибке E6A')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E6A', m.recid,'',;
     	'Отсутствие отчества при d_type!=2,f')
    ENDIF 
   ENDIF 
   
   IF M.E7A == .T. && Алгоритм E7
    IF (!INLIST(people.w,1,2) OR (IsKms(people.sn_pol) AND SUBSTR(people.sn_pol,5,2)!='77' AND (people.w != IIF(VAL(SUBSTR(people.sn_pol,12,2))>50, 1, 2))))
     m.polis = sn_pol
     m.recsproc = 0 
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E7A', m.recid)
*     InsErrorSV(m.mcod,'R', 'E7A', m.recid)
    ENDIF 
   ENDIF 

   IF M.E7A == .T.
    m.sn_pol = people.sn_pol                && Алгоритм E7
    Dtt = CTOD(IIF(VAL(SUBSTR(m.sn_pol,12,2))>50, ;
         PADL(INT(VAL(SUBSTR(m.sn_pol,12,2))-50),2,'0'), ;
         SUBSTR(m.sn_pol,12,2))+'.'+IIF(VAL(SUBSTR(m.sn_pol,14,2))>40, ;
         PADL(INT(VAL(SUBSTR(m.sn_pol,14,2))-40),2,'0')+'.20', ;
         SUBSTR(m.sn_pol,14,2)+'.19')+SUBSTR(m.sn_pol,16,2))
    IF (IsKms(m.sn_pol) AND !INLIST(SUBSTR(m.sn_pol,5,2),'50','51') AND (people.dr != Dtt))
     m.polis = sn_pol
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E7A', m.recid)
*     InsErrorSV(m.mcod, 'R', 'E7A', m.recid)
    ENDIF 
   ENDIF 

   IF M.E8A == .T.
    m.sn_pol = people.sn_pol                && Алгоритм E8
    IF (people.dr=={} OR (dat1-IIF(!EMPTY(people.dr), people.dr, {01.01.1850}))/365.25>120 OR ;
     IIF(!EMPTY(people.dr), people.dr, {01.01.1850}) > m.dat2)
     m.polis = sn_pol
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E8A', m.recid)
*     InsErrorSV(m.mcod,'R', 'E8A', m.recid)
    ENDIF 
   ENDIF 
  
  SELECT c_talon 

RETURN 