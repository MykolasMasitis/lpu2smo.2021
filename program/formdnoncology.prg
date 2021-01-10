PROCEDURE FormDNOncology
 IF MESSAGEBOX('СФОРМИРОВАТЬ ДН_Онкология',4+32,'')=7
  RETURN 
 ENDIF 

 IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\dsp_r.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\dsp_r', 'dsp', 'shar', 'unik')>0
  IF USED('dsp')
   USE IN dsp
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR svod (n_str n(2), dslist m, dsname m,;
  col_28 n(6), col_29 n(6), col_30 n(6), col_31 n(6), col_32 n(6),;
  col_33 n(6), col_34 n(6), col_35 n(6), col_36 n(6), col_37 n(6),;
  col_38 n(6), col_39 n(6), col_40 n(6), col_41 n(6), col_42 n(6),; 
  col_43 n(6), col_44 n(6), col_45 n(6), col_46 n(6), col_47 n(6),;
  col_48 n(6), col_49 n(6), col_50 n(6), col_51 n(6), col_52 n(6),;
  col_53 n(6), col_54 n(6), col_55 n(6), col_56 n(6), col_57 n(6),;
  col_58 n(6), col_59 n(6), col_60 n(6), col_61 n(6), col_62 n(6))
 SELECT svod 
 INDEX on n_str TAG n_str 
 SET ORDER TO n_str
 FOR m.i=11 TO 37 
  DO CASE 
   CASE m.i=11
    m.dslist='С44'
    m.dsname = 'Лица, у которых подтверждено злокачественное новообразование кожи, морфологически определенное как «Базально- клеточный рак» (код МКБ-О-3 8090 –8093), получившие радикальное лечение'
   CASE m.i=12
    m.dslist='С00-С96, исключая Базально-клеточный рак С44,(код МКБ-О-3 8090 –8093), в том числе:'
    m.dsname = 'Лица, с подтвержденным диагнозом ЗНО,'
   CASE m.i=13
    m.dslist='С00'
    m.dsname = 'Губа'
   CASE m.i=14
    m.dslist='С01-С09'
    m.dsname = 'Полость рта'
   CASE m.i=15
    m.dslist='С10-С13'
    m.dsname = 'Глотка'
   CASE m.i=16
    m.dslist='С15'
    m.dsname = 'Пищевод'
   CASE m.i=17
    m.dslist='С16'
    m.dsname = 'Желудок'
   CASE m.i=18
    m.dslist='С18'
    m.dsname = 'Ободочная кишка'
   CASE m.i=19
    m.dslist='С19-С21'
    m.dsname = 'Прямая кишка, ректосиг. соединение, анус'
   CASE m.i=20
    m.dslist='C22'
    m.dsname = 'Печень и внутрипеченочные желчные протоки'
   CASE m.i=21
    m.dslist='C25'
    m.dsname = 'Поджелудочная железа'
   CASE m.i=22
    m.dslist='C32'
    m.dsname = 'Гортань'
   CASE m.i=23
    m.dslist='С33,34'
    m.dsname = 'Трахея, бронхи, легкое'
   CASE m.i=24
    m.dslist='С40,41'
    m.dsname = 'Кости и суставные хрящи'
   CASE m.i=25
    m.dslist='С43'
    m.dsname = 'Меланома кожи'
   CASE m.i=26
    m.dslist='С47,49'
    m.dsname = 'Соединительная и другие мягкие ткани'
   CASE m.i=27
    m.dslist='С50'
    m.dsname = 'Молочная железа'
   CASE m.i=28
    m.dslist='С53'
    m.dsname = 'Шейка матки'
   CASE m.i=29
    m.dslist='С54'
    m.dsname = 'Тело матки'
   CASE m.i=30
    m.dslist='С56'
    m.dsname = 'Яичник'
   CASE m.i=31
    m.dslist='С56'
    m.dsname = 'Предстательная железа'
   CASE m.i=32
    m.dslist='С64'
    m.dsname = 'Почка'
   CASE m.i=33
    m.dslist='С67'
    m.dsname = 'Мочевой пузырь'
   CASE m.i=34
    m.dslist='С73'
    m.dsname = 'Щитовидная железа'
   CASE m.i=35
    m.dslist='С81-86, 88, 90, 96'
    m.dsname = 'Злокачественные лимфомы'
   CASE m.i=36
    m.dslist='С91-95'
    m.dsname = 'Лейкемия'
   CASE m.i=37
    m.dslist='D00-D09'
    m.dsname = 'Лица, с подтвержденным диагнозом ЗНО'
  ENDCASE 

  INSERT INTO svod (n_str, dslist, dsname) VALUES (m.i, m.dslist, m.dsname)
 ENDFOR 

 CREATE CURSOR ppl (sn_pol c(25), ds c(6))
 INDEX on sn_pol+ds TAG unik
 SET ORDER TO unik

 FOR m.t_m=1 TO m.tMonth
  IF !fso.FileExists(m.pBase+'\'+STR(m.tYear,4)+PADL(m.t_m,2,'0')+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(m.pBase+'\'+STR(m.tYear,4)+PADL(m.t_m,2,'0')+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 
  SELECT aisoms

  SCAN 
   m.mcod = mcod 
   IF !fso.FolderExists(m.pBase+'\'+STR(m.tYear,4)+PADL(m.t_m,2,'0')+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(m.pBase+'\'+STR(m.tYear,4)+PADL(m.t_m,2,'0')+'\'+m.mcod+'\people.dbf')
    LOOP 
   ENDIF 
   IF !fso.FileExists(m.pBase+'\'+STR(m.tYear,4)+PADL(m.t_m,2,'0')+'\'+m.mcod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(m.pBase+'\'+STR(m.tYear,4)+PADL(m.t_m,2,'0')+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
    IF USED('people')
     USE IN people 
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   IF OpenFile(m.pBase+'\'+STR(m.tYear,4)+PADL(m.t_m,2,'0')+'\'+m.mcod+'\talon', 'talon', 'shar')>0
    USE IN people 
    IF USED('talon')
     USE IN talon 
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   
   CREATE CURSOR Gosp (c_i c(30))
   SELECT Gosp
   INDEX ON c_i TAG c_i
   SET ORDER TO c_i
  
   WAIT m.mcod+'...' WINDOW NOWAIT 
  
   SELECT talon 
   SET RELATION TO sn_pol INTO people 
   SCAN 
    m.c_i    = c_i
    m.sn_pol = sn_pol
    m.ds     = ds
    m.d_u    = d_u
    DO CASE 
     CASE INLIST(m.ds,'C44')
      m.ds = LEFT(m.ds,3)
      m.row = 11
     CASE BETWEEN(m.ds,'С00','С96') AND m.ds!='C44'
      m.ds = LEFT(m.ds,3)
      m.row = 12
     CASE INLIST(m.ds, 'C00')
      m.ds = LEFT(m.ds,3)
      m.row = 13
     CASE BETWEEN(m.ds, 'C01','C09')
      m.ds = LEFT(m.ds,3)
      m.row = 14
     CASE BETWEEN(m.ds, 'C10','C13')
      m.ds = LEFT(m.ds,3)
      m.row = 15
     CASE INLIST(m.ds, 'C15')
      m.ds = LEFT(m.ds,3)
      m.row = 16
     CASE INLIST(m.ds, 'C16')
      m.ds = LEFT(m.ds,3)
      m.row = 17
     CASE INLIST(m.ds, 'C18')
      m.ds = LEFT(m.ds,3)
      m.row = 18
     CASE BETWEEN(m.ds, 'C19','C21')
      m.ds = LEFT(m.ds,3)
      m.row = 19
     CASE INLIST(m.ds, 'C22')
      m.ds = LEFT(m.ds,3)
      m.row = 20
     CASE INLIST(m.ds, 'C25')
      m.ds = LEFT(m.ds,3)
      m.row = 21
     CASE INLIST(m.ds, 'C32')
      m.ds = LEFT(m.ds,3)
      m.row = 22
     CASE INLIST(m.ds, 'C33','C34')
      m.ds = LEFT(m.ds,3)
      m.row = 23
     CASE INLIST(m.ds, 'C40','C41')
      m.ds = LEFT(m.ds,3)
      m.row = 24
     CASE INLIST(m.ds, 'C43')
      m.ds = LEFT(m.ds,3)
      m.row = 25
     CASE INLIST(m.ds, 'C47', 'C49')
      m.ds = LEFT(m.ds,3)
      m.row = 26
     CASE INLIST(m.ds, 'C50')
      m.ds = LEFT(m.ds,3)
      m.row = 27
     CASE INLIST(m.ds, 'C53')
      m.ds = LEFT(m.ds,3)
      m.row = 28
     CASE INLIST(m.ds,'C54')
      m.ds = LEFT(m.ds,3)
      m.row = 29
     CASE INLIST(m.ds, 'C56')
      m.ds = LEFT(m.ds,3)
      m.row = 30
     CASE INLIST(m.ds, 'C61')
      m.ds = LEFT(m.ds,3)
      m.row = 31
     CASE INLIST(m.ds, 'C64')
      m.ds = LEFT(m.ds,3)
      m.row = 32
     CASE INLIST(m.ds, 'C67')
      m.ds = LEFT(m.ds,3)
      m.row = 33
     CASE INLIST(m.ds, 'C73')
      m.ds = LEFT(m.ds,3)
      m.row = 34
     CASE BETWEEN(m.ds,'С81','C86') OR INLIST(m.ds, 'C88', 'C90', 'C96')
      m.ds = LEFT(m.ds,3)
      m.row = 35
     CASE BETWEEN(m.ds,'С91','C95')
      m.ds = LEFT(m.ds,3)
      m.row = 36
     CASE BETWEEN(m.ds,'D00','D09')
      m.ds = LEFT(m.ds,3)
      m.row = 37
     OTHERWISE 
      LOOP 
    ENDCASE
    
    m.vir = m.sn_pol + m.ds
    IF !SEEK(m.vir, 'dsp')
     LOOP 
    ENDIF  

    m.cod  = cod
    m.rslt = rslt

    m.w    = people.w
    m.dr   = people.dr
    m.adj  = CTOD(STRTRAN(DTOC(m.dr), STR(YEAR(m.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
    m.ages = (YEAR(m.d_u) - YEAR(m.dr)) - IIF(m.adj>0, 1, 0)

    IF INLIST(m.rslt,10,11,12,105,106,205,206,313,405,406,411)
     IF SEEK(m.row, 'svod')
      m.o_col_53 = svod.col_53
      REPLACE col_53 WITH m.o_col_53+1 IN svod
      IF m.w=1
       IF BETWEEN(m.ages,18,65)
        m.o_col_54 = svod.col_54
        REPLACE col_54 WITH m.o_col_54+1 IN svod
       ELSE 
        m.o_col_55 = svod.col_55
        REPLACE col_55 WITH m.o_col_55+1 IN svod
       ENDIF 
      ELSE
       IF BETWEEN(m.ages,18,60)
        m.o_col_56 = svod.col_56
        REPLACE col_56 WITH m.o_col_56+1 IN svod
       ELSE 
        m.o_col_57 = svod.col_57
        REPLACE col_57 WITH m.o_col_57+1 IN svod
       ENDIF 
      ENDIF 
     ENDIF 
    ENDIF 

    DO CASE 
     CASE m.cod = 1195 && посещение врача-онколога
      IF SEEK(m.row, 'svod')
       m.o_col_28 = svod.col_28
       REPLACE col_28 WITH m.o_col_28+1 IN svod
       IF m.w=1
        IF BETWEEN(m.ages,18,65)
         m.o_col_29 = svod.col_29
         REPLACE col_29 WITH m.o_col_29+1 IN svod
        ELSE 
         m.o_col_30 = svod.col_30
         REPLACE col_30 WITH m.o_col_30+1 IN svod
        ENDIF 
       ELSE
        IF BETWEEN(m.ages,18,60)
         m.o_col_31 = svod.col_31
         REPLACE col_31 WITH m.o_col_31+1 IN svod
        ELSE 
         m.o_col_32 = svod.col_32
         REPLACE col_32 WITH m.o_col_32+1 IN svod
        ENDIF 
       ENDIF 

       IF !SEEK(m.vir, 'ppl')
        m.o_col_38 = svod.col_38
        REPLACE col_38 WITH m.o_col_38+1 IN svod
        IF m.w=1
         IF BETWEEN(m.ages,18,65)
          m.o_col_39 = svod.col_39
          REPLACE col_39 WITH m.o_col_39+1 IN svod
         ELSE 
          m.o_col_40 = svod.col_40
          REPLACE col_40 WITH m.o_col_40+1 IN svod
         ENDIF 
        ELSE
         IF BETWEEN(m.ages,18,60)
          m.o_col_41 = svod.col_41
          REPLACE col_41 WITH m.o_col_41+1 IN svod
         ELSE 
          m.o_col_42 = svod.col_42
          REPLACE col_42 WITH m.o_col_42+1 IN svod
         ENDIF 
        ENDIF 

        INSERT INTO ppl FROM MEMVAR 
       ENDIF 

      ENDIF 

     CASE Is02(m.cod) && вызовы скорой помощи
      IF SEEK(m.row, 'svod')
       m.o_col_58 = svod.col_58
       REPLACE col_58 WITH m.o_col_58+1 IN svod
       IF m.w=1
        IF BETWEEN(m.ages,18,65)
         m.o_col_59 = svod.col_59
         REPLACE col_59 WITH m.o_col_59+1 IN svod
        ELSE 
         m.o_col_60 = svod.col_60
         REPLACE col_60 WITH m.o_col_60+1 IN svod
        ENDIF 
       ELSE
        IF BETWEEN(m.ages,18,60)
         m.o_col_61 = svod.col_61
         REPLACE col_61 WITH m.o_col_61+1 IN svod
        ELSE 
         m.o_col_62 = svod.col_62
         REPLACE col_62 WITH m.o_col_62+1 IN svod
        ENDIF 
       ENDIF 
      ENDIF 

     CASE !IsUsl(m.cod) && госпитализация
      IF !SEEK(m.c_i, 'Gosp')
       IF SEEK(m.row, 'svod')
        m.o_col_48 = svod.col_48
        REPLACE col_48 WITH m.o_col_48+1 IN svod
        IF m.w=1
         IF BETWEEN(m.ages,18,65)
          m.o_col_49 = svod.col_49
          REPLACE col_49 WITH m.o_col_49+1 IN svod
         ELSE 
          m.o_col_50 = svod.col_50
          REPLACE col_50 WITH m.o_col_50+1 IN svod
         ENDIF 
        ELSE
         IF BETWEEN(m.ages,18,60)
          m.o_col_51 = svod.col_51
          REPLACE col_51 WITH m.o_col_51+1 IN svod
         ELSE 
          m.o_col_52 = svod.col_52
          REPLACE col_52 WITH m.o_col_52+1 IN svod
         ENDIF 
        ENDIF 
       ENDIF 
       INSERT INTO Gosp FROM MEMVAR 
      ENDIF 

     OTHERWISE 
    ENDCASE 
   
   ENDSCAN 
   SET RELATION OFF INTO people 
   USE IN people 
   USE IN talon 
   USE IN Gosp
   
   SELECT aisoms 
  
   WAIT CLEAR 

  ENDSCAN 
  
  USE IN aisoms 
 ENDFOR 
 
 SELECT svod
 COPY TO &pBase\&gcPeriod\DHOncology WITH cdx 
 m.llResult = X_Report(pTempl+'\DHOncology.xls', pBase+'\'+m.gcperiod+'\DHOncology.xls', .T.)
 
 USE 
 
 USE IN ppl 
 USE IN dsp 

RETURN 