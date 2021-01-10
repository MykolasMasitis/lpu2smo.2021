PROCEDURE ExpExpI3c
 IF !m.IsServer
  MESSAGEBOX('дюммши бхд щйяонпрю бнглнфем'+CHR(13)+CHR(10)+'рнкэйн б пефхле "яепбепю"!',0+64,'')
  RETURN 
 ENDIF 
 IF MESSAGEBOX(CHR(13)+CHR(10)+'бш унрхре гюцпсгхрэ пегскэрюрш щйяоепрхг?'+CHR(13)+CHR(10),4+32,'хмцняярпюу')=7
  RETURN 
 ENDIF 

 IF !fso.FolderExists(pExpImp)
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер дхпейрнпхъ щйяонпрю-хлонпрю'+CHR(13)+CHR(10)+pExpImp+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pExpImp+'\catalog.dbf')
  MESSAGEBOX('тюик '+pExpImp+'\catalog.dbf'+'ме мюидем!',0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pExpImp+'\rss.dbf')
  MESSAGEBOX('тюик '+pExpImp+'\rss.dbf'+'ме мюидем!',0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pExpImp+'\svacts.dbf')
  MESSAGEBOX('тюик '+pExpImp+'\svacts.dbf'+'ме мюидем!',0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pExpImp+'\ssacts.dbf')
  MESSAGEBOX('тюик '+pExpImp+'\ssacts.dbf'+'ме мюидем!',0+64,'')
  RETURN 
 ENDIF 


 IF !fso.FolderExists(pMee+'\REQUESTS')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер дхпейрнпхъ гюопнянб'+CHR(13)+CHR(10)+pMee+'\REQUESTS'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pMee+'\RSS')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер дхпейрнпхъ пееярпнб'+CHR(13)+CHR(10)+pMee+'\RSS'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pMee+'\SVACTS')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер дхпейрнпхъ ябндмшу юйрнб'+CHR(13)+CHR(10)+pMee+'\SVACTS'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pMee+'\SSACTS')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер дхпейрнпхъ оепянмюкэмшу юйрнб'+CHR(13)+CHR(10)+pMee+'\SSCTS'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pMee+'\REQUESTS\catalog.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер йюрюкнц гюопнянб'+CHR(13)+CHR(10)+pMee+'\REQUESTS\catalog.dbf'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pMee+'\RSS\rss.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер йюрюкнц пееярпнб'+CHR(13)+CHR(10)+pMee+'\RSS\rss.dbf'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pMee+'\SVACTS\svacts.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер йюрюкнц ябндмшу юйрнб'+CHR(13)+CHR(10)+pMee+'\SVACTS\svacts.dbf'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pMee+'\SSACTS\ssacts.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер йюрюкнц оепянмюкэмшу юйрнб'+CHR(13)+CHR(10)+pMee+'\SSACTS\ssacts.dbf'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pMee+'\SVACTS\svacts', 'svacts', 'shar', 'recid')>0
  IF USED('svacts')
   USE IN svacts
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pMee+'\SVACTS\moves', 'svmoves', 'shar', 'actid')>0
  USE IN svacts
  IF USED('svmoves')
   USE IN svmoves
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pMee+'\SSACTS\ssacts', 'ssacts', 'shar', 'recid')>0
  USE IN svacts
  USE IN svmoves
  IF USED('ssacts')
   USE IN ssacts
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pMee+'\REQUESTS\catalog', 'rqst', 'shar', 'recid')>0
  USE IN svacts
  USE IN ssacts
  USE IN svmoves
  IF USED('rqst')
   USE IN rqst
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pMee+'\RSS\rss', 'rss', 'shar', 'recid')>0
  USE IN svacts
  USE IN svmoves
  USE IN ssacts
  USE IN rqst
  IF USED('rss')
   USE IN rss
  ENDIF 
  RETURN 
 ENDIF 
 
 IF fso.FileExists(pExpImp+'\catalog.dbf')
  IF OpenFile(pExpImp+'\catalog.dbf', 'imprqst', 'shar')>0
   IF USED('imprqst')
    USE IN imprqst
   ENDIF 
  ELSE 
   WAIT "йнохпсч гюопняш..." WINDOW NOWAIT 
   SELECT imprqst
   SCAN 
    SCATTER MEMVAR 
    m.rqfile = PADL(m.recid,6,'0')
    IF fso.FileExists(pExpImp+'\'+m.rqfile+'.dbf') AND fso.FileExists(pExpImp+'\'+m.rqfile+'.cdx')
     IF !fso.FileExists(pMee+'\REQUESTS\'+m.rqfile+'.dbf')
      fso.CopyFile(pExpImp+'\'+m.rqfile+'.dbf', pMee+'\REQUESTS\'+m.rqfile+'.dbf')
      fso.CopyFile(pExpImp+'\'+m.rqfile+'.cdx', pMee+'\REQUESTS\'+m.rqfile+'.cdx')
      
      IF OpenFile(pMee+'\REQUESTS\'+m.rqfile+'.dbf', 'numrq', 'excl')>0
       IF USED('numrq')
        USE IN numrq
       ENDIF 
       SELECT imprqst
      ELSE 
       SELECT numrq
       DELETE TAG ALL 
       INDEX on sn_pol TAG sn_pol
       SET ORDER TO sn_pol
       
       SCAN 
        m.n_akt = n_akt
        m.d_akt = d_akt
        m.t_akt = t_akt
        m.r_id  = INT(VAL(SUBSTR(m.n_akt,10,6)))

        IF !EMPTY(m.n_akt)
         DO CASE 
          CASE m.t_akt = 'SS'
           m.sn_pol = sn_pol
           m.fio    = ALLTRIM(fio)
           m.fam    = SUBSTR(m.fio, 1, AT(SPACE(1), m.fio, 1)-1)
           m.im     = SUBSTR(m.fio, AT(SPACE(1), m.fio, 1)+1, AT(SPACE(1), m.fio, 2)-AT(SPACE(1), m.fio, 1))
           m.ot     = SUBSTR(m.fio, AT(SPACE(1), m.fio, 2)+1)
           IF m.r_id>0 AND !SEEK(m.r_id, 'ssacts')
            *INSERT INTO ssacts (recid, n_akt, mcod, lpu_id, period, e_period, smoexp, docexp, codexp, reason, actdate, sn_pol, fam, im, ot, qr, status) VALUES ;
            	(m.r_id, m.n_akt, m.mcod, m.lpu_id, m.period, m.e_period, m.smoexp, m.supexp, VAL(m.et), m.rs, m.d_akt, m.sn_pol, m.fam, m.im, m.ot, .T., '1')
           ENDIF 
           
          CASE m.t_akt = 'SV'
           IF m.r_id>0 AND !SEEK(m.r_id, 'svacts')
            *INSERT INTO svacts (recid, n_akt, mcod, lpu_id, period, e_period, smoexp, docexp, et, reason, codexp, actdate, qr, status) VALUES ;
            	(m.r_id, m.n_akt, m.mcod, m.lpu_id, m.period, m.e_period, m.smoexp, m.supexp, m.et, m.rs, VAL(m.et), m.d_akt, .T., '1')
           ENDIF 
          OTHERWISE 
         ENDCASE 
        ENDIF 

       ENDSCAN 
       USE IN numrq

       IF !SEEK(m.recid, 'rqst')
        INSERT INTO rqst FROM MEMVAR 
       ENDIF 

       SELECT imprqst
      ENDIF 

     ELSE 

      IF OpenFile(pMee+'\REQUESTS\'+m.rqfile+'.dbf', 'numrq', 'excl')>0
       IF USED('numrq')
        USE IN numrq
       ENDIF 
      ELSE 
       SELECT numrq
       DELETE TAG ALL 
       INDEX on sn_pol TAG sn_pol
       SET ORDER TO sn_pol
       IF OpenFile(pExpImp+'\'+m.rqfile+'.dbf', 'impnumrq', 'excl')>0
        USE IN numrq
        IF USED('impnumrq')
         USE IN impnumrq
        ENDIF 
       ELSE 
        SELECT impnumrq
        DELETE TAG ALL 
        INDEX on sn_pol TAG sn_pol
        SET ORDER TO sn_pol
        SCAN 
         **
         m.n_akt = n_akt
         m.d_akt = d_akt
         m.t_akt = t_akt
         m.r_id  = INT(VAL(SUBSTR(m.n_akt,10,6)))

         IF !EMPTY(m.n_akt)
          DO CASE 
           CASE m.t_akt = 'SS'
            IF m.r_id>0 AND !SEEK(m.r_id, 'ssacts')
             *INSERT INTO ssacts (recid, n_akt, mcod, lpu_id, period, e_period) VALUES ;
             	(m.r_id, m.n_akt, m.mcod, m.lpu_id, m.period, m.e_period)
            ENDIF 
           CASE m.t_akt = 'SV'
            IF m.r_id>0 AND !SEEK(m.r_id, 'svacts')
             *INSERT INTO svacts (recid, n_akt, mcod, lpu_id, period, e_period) VALUES ;
             	(m.r_id, m.n_akt, m.mcod, m.lpu_id, m.period, m.e_period)
            ENDIF 
           OTHERWISE 
          ENDCASE 
         ENDIF 
         **
         m.snp = sn_pol
         IF SEEK(m.snp, 'numrq')
          IF EMPTY(numrq.n_akt)
           m.n_akt = n_akt
           m.d_akt = d_akt
           m.t_akt = t_akt
           REPLACE n_akt WITH m.n_akt, d_akt WITH m.d_akt, t_akt WITH m.t_akt IN numrq
          ENDIF 
         ENDIF 
        ENDSCAN 
        USE IN impnumrq
        USE IN numrq
       ENDIF 
      ENDIF 
     ENDIF 

     IF !SEEK(m.recid, 'rqst')
      INSERT INTO rqst FROM MEMVAR 
     ENDIF 

    ENDIF 
   ENDSCAN 
   USE IN imprqst
   WAIT CLEAR 
  ENDIF 
 ENDIF 

 IF fso.FileExists(pExpImp+'\rss.dbf')
  IF OpenFile(pExpImp+'\rss.dbf', 'imprss', 'shar')>0
   IF USED('imprss')
    USE IN imprss
   ENDIF 
  ELSE 
   WAIT "йнохпсч пееярп юйрнб..." WINDOW NOWAIT 
   SELECT imprss
   SCAN 
    SCATTER MEMVAR 
    IF !SEEK(m.recid, 'rss')
     INSERT INTO rss FROM MEMVAR 
    ENDIF 
   ENDSCAN 
   USE IN imprss
   WAIT CLEAR 
  ENDIF 
 ENDIF 

 IF fso.FileExists(pExpImp+'\svacts.dbf')
  IF OpenFile(pExpImp+'\svacts.dbf', 'impsv', 'shar')>0
   IF USED('impsv')
    USE IN impsv
   ENDIF 
  ELSE 
   WAIT "йнохпсч ябндмше юйрш..." WINDOW NOWAIT 
   SELECT impsv
   SCAN 
    SCATTER MEMVAR 
    m.e_period = STR(YEAR(DATE()),4)+PADL(MONTH(DATE()),2,'0')
    m.status   = '1'
    IF !SEEK(m.recid, 'svacts')
     INSERT INTO svacts FROM MEMVAR 
     INSERT INTO svmoves (actid, et, usr, dat) VALUES (m.recid, m.et, m.gcUser, DATETIME())
    ENDIF 
   ENDSCAN 
   USE IN impsv
   WAIT CLEAR 
  ENDIF 
 ENDIF 
 
 IF fso.FileExists(pExpImp+'\ssacts.dbf')
  IF OpenFile(pExpImp+'\ssacts.dbf', 'impss', 'shar')>0
   IF USED('impss')
    USE IN impss
   ENDIF 
  ELSE 
   WAIT "йнохпсч хмдхбхдсюкэмше юйрш..." WINDOW NOWAIT 
   SELECT impss
   SCAN 
    SCATTER MEMVAR 
    m.e_period = STR(YEAR(DATE()),4)+PADL(MONTH(DATE()),2,'0')
    m.status   = '1'
    IF !SEEK(m.recid, 'ssacts')
     INSERT INTO ssacts FROM MEMVAR 
    ENDIF 
   ENDSCAN 
   USE IN impss
   WAIT CLEAR 
  ENDIF 
 ENDIF 
 
* USE IN svacts
 USE IN ssacts
 USE IN rqst
 USE IN rss

 USE IN svacts
 USE IN svmoves

 WAIT CLEAR 

 MESSAGEBOX(CHR(13)+CHR(10)+'напюанрюмн '+ALLTRIM(STR(m.nGoodFiles))+' тюикнб'+;
 CHR(13)+CHR(10),0+64,'')

RETURN 