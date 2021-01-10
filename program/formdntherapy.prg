PROCEDURE FormDNTherapy
 IF MESSAGEBOX('СФОРМИРОВАТЬ ДН_Терапия',4+32,'')=7
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
    m.dslist='I20.1,I20.8,I20.9,I25.0,I25.1,I25.2,I25.5,I25.6,I25.8,I25.9'
    m.dsname = 'Стабильная ишемическая болезнь сердца (за исключением следующих заболеваний или состояний, по поводу которых осуществляется диспансерное наблюдение врачом-кардиологом:'
    m.dsname = m.dsname + 'стенокардия III - IV ФК в трудоспособном возрасте;'
    m.dsname = m.dsname + 'перенесенный инфаркт миокарда и его осложнений в течение 12 месяцев после оказания медицинской помощи в стационарных условиях медицинских организаций;'
    m.dsname = m.dsname + 'период после оказания высокотехнологичных методов лечения, включая кардиохирургические вмешательства в течение 12 месяцев после оказания медицинской помощи в стационарных условиях медицинских организаций)'
   CASE m.i=12
    m.dslist='I10,I11,I12,I13,I15'
    m.dsname = 'Артериальная гипертония 1 - 3 степени, за исключением резистентной артериальной гипертонии'
   CASE m.i=13
    m.dslist='I50.0,I50.1,I50.9'
    m.dsname = 'Хроническая сердечная недостаточность I - III ФК по NYHA, но не выше стадии 2а'
   CASE m.i=14
    m.dslist='I48'
    m.dsname = 'Фибрилляция и (или) трепетание предсердий'
   CASE m.i=15
    m.dslist='I47'
    m.dsname = 'Предсердная и желудочковая экстрасистолия, наджелудочковые и желудочковые тахикардии на фоне эффективной профилактической антиаритмической терапии'
   CASE m.i=16
    m.dslist='I65.2'
    m.dsname = 'Стеноз внутренней сонной артерии от 40 до 70%'
   CASE m.i=17
    m.dslist='R73.0,R73.9'
    m.dsname = 'Предиабет'
   CASE m.i=18
    m.dslist='E11'
    m.dsname = 'Сахарный диабет 2 типа'
   CASE m.i=19
    m.dslist='I69.0,I69.1,I69.2,I69.3,I69.4,I67.8'
    m.dsname = 'Последствия перенесенных острых нарушений мозгового кровообращения'
   CASE m.i=20
    m.dslist='E78'
    m.dsname = 'Гиперхолестеринемия (при уровне общего холестерина более 8,0 ммоль/л)'
   CASE m.i=21
    m.dslist='K20'
    m.dsname = 'Эзофагит (эозинофильный, химический, лекарственный)'
   CASE m.i=22
    m.dslist='K21.0'
    m.dsname = 'Гастроэзофагеальный рефлюкс с эзофагитом (без цилиндроклеточной метаплазии - без пищевода Баррета)'
   CASE m.i=23
    m.dslist='K21.0'
    m.dsname = 'Гастроэзофагеальный рефлюкс с эзофагитом и цилиндроклеточной метаплазией - пищевод Барретта'
   CASE m.i=24
    m.dslist='K25'
    m.dsname = 'Язвенная болезнь желудка'
   CASE m.i=25
    m.dslist='K26'
    m.dsname = 'Язвенная болезнь двенадцатиперстной кишки'
   CASE m.i=26
    m.dslist='K29.4,K29.5'
    m.dsname = 'Хронический атрофический фундальный и мультифокальный гастрит'
   CASE m.i=27
    m.dslist='K31.7'
    m.dsname = 'Полипы (полипоз) желудка'
   CASE m.i=28
    m.dslist='K86'
    m.dsname = 'Хронический панкреатит с внешнесекреторной недостаточностью'
   CASE m.i=29
    m.dslist='J44.0,J44.8,J44.9'
    m.dsname = 'Хроническая обструктивная болезнь легких'
   CASE m.i=30
    m.dslist='J47.0'
    m.dsname = 'Бронхоэктатическая болезнь'
   CASE m.i=31
    m.dslist='J45.0,J45.1,J45.8,J45.9'
    m.dsname = 'Бронхиальная астма'
   CASE m.i=32
    m.dslist='J12,J13,J14'
    m.dsname = 'Состояние после перенесенной пневмонии'
   CASE m.i=33
    m.dslist='J84.1'
    m.dsname = 'Интерстициальные заболевания легких'
   CASE m.i=34
    m.dslist='N18.1'
    m.dsname = 'Пациенты, перенесшие острую почечную недостаточность, в стабильном состоянии, с хронической почечной недостаточностью 1 стадии'
   CASE m.i=35
    m.dslist='N18.1'
    m.dsname = 'Пациенты, страдающие хронической болезнью почек (независимо от ее причины и стадии), в стабильном состоянии с хронической почечной недостаточностью 1 стадии'
   CASE m.i=36
    m.dslist='N18.9'
    m.dsname = 'Пациенты, относящиеся к группам риска поражения почек'
   CASE m.i=37
    m.dslist='M81.5'
    m.dsname = 'Остеопороз первичный'
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
     CASE INLIST(m.ds,'I20.1', 'I20.8', 'I20.9', 'I25.0', 'I25.1', 'I25.2', 'I25.5', 'I25.6', 'I25.8', 'I25.9')
      m.ds = LEFT(m.ds,5)
      m.row = 11
     CASE INLIST(m.ds, 'I10', 'I11', 'I12', 'I13', 'I15')
      m.ds = LEFT(m.ds,3)
      m.row = 12
     CASE INLIST(m.ds, 'I50.0', 'I50.1', 'I50.9')
      m.ds = LEFT(m.ds,5)
      m.row = 13
     CASE INLIST(m.ds, 'I48')
      m.ds = LEFT(m.ds,3)
      m.row = 14
     CASE INLIST(m.ds, 'I47')
      m.ds = LEFT(m.ds,3)
      m.row = 15
     CASE INLIST(m.ds, 'I65.2')
      m.ds = LEFT(m.ds,5)
      m.row = 16
     CASE INLIST(m.ds, 'R73.0', 'R73.9')
      m.ds = LEFT(m.ds,5)
      m.row = 17
     CASE INLIST(m.ds, 'E11')
      m.ds = LEFT(m.ds,3)
      m.row = 18
     CASE INLIST(m.ds, 'I69.0', 'I69.1', 'I69.2', 'I69.3', 'I69.4', 'I67.8')
      m.ds = LEFT(m.ds,5)
      m.row = 19
     CASE INLIST(m.ds, 'E78')
      m.ds = LEFT(m.ds,3)
      m.row = 20
     CASE INLIST(m.ds, 'K20')
      m.ds = LEFT(m.ds,3)
      m.row = 21
     CASE INLIST(m.ds, 'K21.0')
      m.ds = LEFT(m.ds,5)
      m.row = 22
     CASE INLIST(m.ds, 'K21.0')
      m.ds = LEFT(m.ds,5)
      m.row = 23
     CASE INLIST(m.ds, 'K25')
      m.ds = LEFT(m.ds,3)
      m.row = 24
     CASE INLIST(m.ds, 'K26')
      m.ds = LEFT(m.ds,3)
      m.row = 25
     CASE INLIST(m.ds, 'K29.4', 'K29.5')
      m.ds = LEFT(m.ds,5)
      m.row = 26
     CASE INLIST(m.ds, 'K31.7')
      m.ds = LEFT(m.ds,5)
      m.row = 27
     CASE INLIST(m.ds, 'K86')
      m.ds = LEFT(m.ds,3)
      m.row = 28
     CASE INLIST(m.ds,'J44.0', 'J44.8', 'J44.9')
      m.ds = LEFT(m.ds,5)
      m.row = 29
     CASE INLIST(m.ds, 'J47.0')
      m.ds = LEFT(m.ds,5)
      m.row = 30
     CASE INLIST(m.ds, 'J45.0', 'J45.1', 'J45.8', 'J45.9')
      m.ds = LEFT(m.ds,5)
      m.row = 31
     CASE INLIST(m.ds, 'J12', 'J13', 'J14')
      m.ds = LEFT(m.ds,3)
      m.row = 32
     CASE INLIST(m.ds, 'J84.1')
      m.ds = LEFT(m.ds,5)
      m.row = 33
     CASE INLIST(m.ds, 'N18.1')
      m.ds = LEFT(m.ds,5)
      m.row = 34
     CASE INLIST(m.ds, 'N18.1')
      m.ds = LEFT(m.ds,5)
      m.row = 35
     CASE INLIST(m.ds, 'N18.9')
      m.ds = LEFT(m.ds,5)
      m.row = 36
     CASE INLIST(m.ds, 'M81.5')
      m.ds = LEFT(m.ds,5)
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
     CASE m.cod = 1015 && посещение врача-терапевта
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
      
     CASE m.cod = 1016 && посещение врача-терапевта на дому
      IF SEEK(m.row, 'svod')
       m.o_col_33 = svod.col_33
       REPLACE col_33 WITH m.o_col_33+1 IN svod
       IF m.w=1
        IF BETWEEN(m.ages,18,65)
         m.o_col_34 = svod.col_34
         REPLACE col_34 WITH m.o_col_34+1 IN svod
        ELSE 
         m.o_col_35 = svod.col_35
         REPLACE col_35 WITH m.o_col_35+1 IN svod
        ENDIF 
       ELSE
        IF BETWEEN(m.ages,18,60)
         m.o_col_36 = svod.col_36
         REPLACE col_36 WITH m.o_col_36+1 IN svod
        ELSE 
         m.o_col_37 = svod.col_37
         REPLACE col_37 WITH m.o_col_37+1 IN svod
        ENDIF 
       ENDIF 

       IF !SEEK(m.vir, 'ppl')
        m.o_col_43 = svod.col_43
        REPLACE col_43 WITH m.o_col_43+1 IN svod
        IF m.w=1
         IF BETWEEN(m.ages,18,65)
          m.o_col_44 = svod.col_44
          REPLACE col_44 WITH m.o_col_44+1 IN svod
         ELSE 
          m.o_col_45 = svod.col_45
          REPLACE col_45 WITH m.o_col_45+1 IN svod
         ENDIF 
        ELSE
         IF BETWEEN(m.ages,18,60)
          m.o_col_46 = svod.col_46
          REPLACE col_46 WITH m.o_col_46+1 IN svod
         ELSE 
          m.o_col_47 = svod.col_47
          REPLACE col_47 WITH m.o_col_47+1 IN svod
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
 COPY TO &pBase\&gcPeriod\DHTherapy WITH cdx 
 m.llResult = X_Report(pTempl+'\DHTherapy.xls', pBase+'\'+m.gcperiod+'\DHTherapy.xls', .T.)
 
 USE 
 
 USE IN ppl 
 USE IN dsp 

RETURN 