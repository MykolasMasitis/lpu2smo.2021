PROCEDURE ProfRepT1
 IF MESSAGEBOX('СВЕДЕНИЯ ОБ ОБЪЕМАХ И СТОИМОСТИ'+CHR(13)+CHR(10)+;
 	'ДИСПАНСЕРИЗАЦИИ ВЗРОСЛОГО НАСЕЛЕНИЯ?'+CHR(13)+CHR(10)+'',4+64, 'ТАБЛИЦА 1')=7
 	RETURN 
 ENDIF 
 IF !fso.FileExists(pTempl+'\Prof_t1.xls')
  MESSAGEBOX('ШАБЛОН '+UPPER('Prof_t1.xls')+' НЕ НАЙДЕН!',4+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pcommon+'\dspcodes.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ФАЙЛ DSPCODES.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\dsp.dbf')
  MESSAGEBOX('ФАЙЛ dsp.dbf ЗА ВЫБРАННЫЙ ПЕРИООД НЕ СФОРМИРОВАН!', 4+64, '')
  RETURN 
 ENDIF 
 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\dsp', 'dsp', 'shar')>0
  IF USED('dsp')
   USE IN dsp 
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\dspcodes', 'dspcodes', 'shar', 'cod')>0
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  IF USED('dsp')
   USE IN dsp
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR curdsp (sn_pol c(25))
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 DIMENSION dimdsp(11,17)
 dimdsp = 0
 DIMENSION dimdata(3,3)
 dimdata = 0

 SELECT dsp 
 SCAN 
  m.cod    = cod
  m.w      = w
  m.ages   = ages
  m.sn_pol = sn_pol
  m.mcod   = mcod

  IF !SEEK(m.cod, 'dspcodes')
   LOOP
  ENDIF 
  m.tipofcod = dspcodes.tip
  IF m.tipofcod!=1
   LOOP 
  ENDIF 
  IF m.ages<18
   LOOP 
  ENDIF 
  IF m.w=1 AND m.cod = 25204 AND !INLIST(m.ages,49,53,55,59,61,65,67,71,73)
   LOOP 
  ENDIF 
  IF m.w=2 AND !INLIST(m.ages,49,53,55,59,61,65,67,71,73) AND INLIST(m.cod, 25204)
   LOOP 
  ENDIF 
  IF m.w=2 AND !INLIST(m.ages,50,52,56,58,62,64,68,70) AND INLIST(m.cod, 35401)
   LOOP 
  ENDIF 
  
  =incdimdsp(1)
  IF m.w=1 && вторая строка, мужчины
   =incdimdsp(2)
  ENDIF 
  IF m.w=2 && третья строка, женщины
   =incdimdsp(3)
  ENDIF 

  IF (m.w=1 AND BETWEEN(m.ages, 18, 60)) OR (m.w=2 AND BETWEEN(m.ages, 18, 55))
   =incdimdsp(4)
  ENDIF 
  IF m.w=1 AND BETWEEN(m.ages, 18, 60) && пятая строка, мужчины
   =incdimdsp(5)
  ENDIF 
  IF m.w=2 AND BETWEEN(m.ages, 18, 55) && шестая строка, женщины
   =incdimdsp(6)
  ENDIF 
  IF (m.w=1 AND m.ages>60) OR (m.w=2 AND m.ages>55)
   =incdimdsp(7)
  ENDIF 
  IF m.w=1 and m.ages>60 && восьмая строка, мужчины
   =incdimdsp(8)
  ENDIF 
  IF m.w=2 AND m.ages>55 && девятая строка, женщины
   =incdimdsp(9)
  ENDIF 

  IF m.ages=65
   =incdimdsp(10)
  ENDIF 
  IF m.ages>65
   =incdimdsp(11)
  ENDIF 

 ENDSCAN 


 USE IN curdsp
 USE IN dsp 
 USE IN dspcodes
 
 m.nstr = 1
 dimdata(1,2) = TRANSFORM(dimdsp(m.nstr,6)-dimdsp(m.nstr,9), '99999') && пациентов
 dimdata(1,3) = TRANSFORM(dimdsp(m.nstr,7), '999999999.99')           && сумма

 m.nstr = 10
 dimdata(2,2) = TRANSFORM(dimdsp(m.nstr,6)-dimdsp(m.nstr,9), '99999') && пациентов
 dimdata(2,3) = TRANSFORM(dimdsp(m.nstr,7), '999999999.99')           && сумма

 m.nstr = 11
 dimdata(3,2) = TRANSFORM(dimdsp(m.nstr,6)-dimdsp(m.nstr,9), '99999') && пациентов
 dimdata(3,3) = TRANSFORM(dimdsp(m.nstr,7), '999999999.99')           && сумма

 m.lcTmpName = pTempl+'\Prof_t1.xls'
 m.lcRepName = pBase+'\'+m.gcPeriod+'\Prof_t1.xls'
 m.IsVisible = .T.
 
 CREATE CURSOR curdata (recid i)
 m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)
 USE IN curdata
 
RETURN 

FUNCTION incdimdsp(nstr)
 PRIVATE nstr
  dimdsp(m.nstr,1) = dimdsp(m.nstr,1) + 1
  dimdsp(m.nstr,2) = dimdsp(m.nstr,2) + s_all

  IF !EMPTY(k_u2)
   dimdsp(m.nstr,3) = dimdsp(m.nstr,3) + 1
   dimdsp(m.nstr,4) = dimdsp(m.nstr,4) + k_u2
   dimdsp(m.nstr,5) = dimdsp(m.nstr,5) + s_all2
  ENDIF 

  IF EMPTY(er)
   dimdsp(m.nstr,6) = dimdsp(m.nstr,6) + 1
   dimdsp(m.nstr,7) = dimdsp(m.nstr,7) + s_all
   IF !EMPTY(k_u2)
    dimdsp(m.nstr,9) = dimdsp(m.nstr,9) + 1
    dimdsp(m.nstr,10) = dimdsp(m.nstr,10) + k_u2ok
    dimdsp(m.nstr,11) = dimdsp(m.nstr,11) + s_all2ok
   ENDIF 

   DO CASE 
    CASE INLIST(rslt,352,353,357,358)
     dimdsp(m.nstr,12) = dimdsp(m.nstr,12) + 1
    CASE INLIST(rslt,317,352)
     dimdsp(m.nstr,13) = dimdsp(m.nstr,13) + 1
    CASE INLIST(rslt,318,353)
     dimdsp(m.nstr,14) = dimdsp(m.nstr,14) + 1
    CASE INLIST(rslt,319,354)
     dimdsp(m.nstr,15) = dimdsp(m.nstr,15) + 1
    CASE INLIST(rslt,355,357)
     dimdsp(m.nstr,16) = dimdsp(m.nstr,16) + 1
    CASE INLIST(rslt,356,358)
     dimdsp(m.nstr,17) = dimdsp(m.nstr,17) + 1
   ENDCASE 
  ENDIF 

RETURN 
