PROCEDURE cfgbase
LOCAL m.StartPath, m.HomePath
m.StartPath = SYS(5)+SYS(2003)
m.HomePath  = SUBSTR(SYS(5)+SYS(2003), 1, rat('\',SYS(5)+SYS(2003)))

*IF FILE('lpu2smo.cfg')
IF fso.FileExists(m.StartPath+'\lpu2smo.cfg')
 USE lpu2smo.cfg IN 0 ALIAS cnfg EXCLUSIVE 
 *SELECT cnfg
 *IF !FILE('lpu2smo.ocx')
 * INDEX ON qcod FOR IsAisValid() TAG qcod OF lpu2smo.ocx  
 *ENDIF 
 *IF !FILE('lpu2smo.cdx')
 * INDEX ON qcod FOR IsAisValid() TAG qcod
 *ENDIF 

 IF UPPER(FIELD('T_START')) !=  'T_START'
  ALTER TABLE cnfg ADD COLUMN t_start n(2)
  UPDATE cnfg SET t_start=3
 ENDIF 

 IF UPPER(FIELD('T_ROBOT')) !=  'T_ROBOT'
  ALTER TABLE cnfg ADD COLUMN t_robot n(3)
  UPDATE cnfg SET t_robot=20
 ENDIF 

 IF UPPER(FIELD('ISPARALLEL')) !=  'ISPARALLEL'
  ALTER TABLE cnfg ADD COLUMN ISPARALLEL L
 ENDIF 

 IF UPPER(FIELD('ISTESTMODE')) !=  'ISTESTMODE'
  ALTER TABLE cnfg ADD COLUMN ISTESTMODE L
 ENDIF 

 IF UPPER(FIELD('PMEE')) !=  'PMEE'
  ALTER TABLE cnfg ADD COLUMN pmee c(100)
 ENDIF 
 
ELSE 
 IF MESSAGEBOX('ОТСУТСТВУЕТ КОНФИГУРАЦИОННЫЙ ФАЙЛ LPU2SMO.CFG!' + CHR(13) + 'СОЗДАТЬ?',4+32,'ВНИМАНИЕ!') = 6
	CREATE TABLE lpu2smo ;
		(qcod c(2), tyear n(4), tmonth n(4), tdat1 d, tdat2 d, User c(6), Frmt c(3), ;
		 paisoms c(100), parc c(100), pbase c(100), pbin c(100), pcommon c(100), ;
		 pout c(100), plocal c(100), pexpimp c(100), ptempl c(100), ;
		 pmee c(100), IsUsrDir L, ISSERVER l, comport n(2), t_robot n(3), ISPARALLEL L, ISTESTMODE L)
	SELECT lpu2smo
	USE 
	RENAME  lpu2smo.dbf TO lpu2smo.cfg
	INSERT INTO lpu2smo.cfg (qcod, tyear, tmonth, tdat1, tdat2, user, Frmt, paisoms, parc, pbase, pbin, pcommon, ;
	                         pout, plocal, pexpimp, ptempl, pmee, IsUsrDir, ISSERVER, comport, t_robot, ISPARALLEL, ISTESTMODE) ;
		VALUES ('I3', YEAR(DATE()), MONTH(DATE()), CTOD('01.'+PADL(MONTH(DATE()),2,'0')+'.'+PADL(YEAR(DATE()),4,'0')), ;
		GOMONTH(CTOD('01.'+PADL(MONTH(DATE()),2,'0')+'.'+PADL(YEAR(DATE()),4,'0')),1)-1, 'OMS', 'PDF', ;
		 m.HomePath+'AISOMS', m.HomePath+'ARC', m.HomePath+'BASE', m.StartPath, m.HomePath+'COMMON', ;
		 m.HomePath+'OUT', m.HomePath+'LOCAL', m.HomePath+'EXCHANGE.DIR',  m.HomePath+'MAIL', m.HomePath+'TEMPLATES', ;
		 m.HomePath+'TRASH', m.HomePath+'DOUBLES', m.HomePath+'MEE', '', '', .F., .F., 3, 20, .f.)
	SELECT lpu2smo
 ELSE
	MESSAGEBOX('ПРОДОЛЖЕНИЕ РАБОТЫ НЕВОЗМОЖНО!',0+16,'ВНИМАНИЕ!')
	RETURN -1
 ENDIF 
ENDIF 

m.qcod      = ALLTRIM(qcod)
m.pAisOms   = ALLTRIM(paisoms)
m.parc      = ALLTRIM(parc)
m.pbin      = ALLTRIM(pbin)
m.pbase     = ALLTRIM(pbase)
m.pcommon   = ALLTRIM(pcommon)
m.pout      = ALLTRIM(pout)
m.plocal    = ALLTRIM(plocal)
m.pexpimp   = ALLTRIM(pexpimp)
m.ptempl    = ALLTRIM(ptempl)
m.pmee      = ALLTRIM(pmee)
m.tyear     = tyear
m.IsServer  = ISSERVER
m.ComPort   = comport
m.tmonth    = tmonth
m.tdat1     = tdat1
m.tdat2     = tdat2
m.gcUser    = ALLTRIM(User)
m.gcFormat  = UPPER(Frmt)
m.IsUsrDir  = IsUsrDir
m.t_robot   = t_robot
m.t_start   = t_start
m.ISPARALLEL = ISPARALLEL
m.ISTESTMODE = ISTESTMODE

*USE IN cnfg

*IF m.HomePath+'BIN' != m.pBin
IF m.StartPath != m.pBin
	MESSAGEBOX("Директория запуска программы"+CHR(13)+;
	"отличается от директории по умолчанию."+CHR(13)+;
	"Сделать директорию запуска новой директорией по умолчанию?",4+32,"Внимание!")
	
	*m.paisoms   = m.HomePath + 'AISOMS'
	m.parc      = m.HomePath + 'ARC'
	m.pbin      = m.HomePath + 'BIN'
	m.pbase     = m.HomePath + 'BASE'
	*m.pcommon   = m.HomePath + 'COMMON'
	m.pout      = m.HomePath + 'OUT'
	m.plocal    = m.HomePath + 'LOCAL'
	m.pexpimp   = m.HomePath + 'EXCHANGE.DIR'
	m.ptempl    = m.HomePath + 'TEMPLATES'
	m.pmee      = m.HomePath + 'MEE'
    m.IsUsrDir  = .F.
    m.comport  = 3
	REPLACE parc WITH m.parc, pbin WITH m.pbin, pbase WITH m.pbase, pout WITH m.pout ,;
		plocal WITH m.plocal, pexpimp WITH m.pexpimp,;
		ptempl WITH m.ptempl, ;
		pmee WITH m.pmee, IsUsrDir WITH m.IsUsrDir
	
ENDIF

m.dirchk = FDATE(m.pbin+'\lpu2smo.exe')

IF USED('lpu2smo')
 USE IN lpu2smo
ENDIF 
IF USED('cnfg')
 USE IN cnfg
ENDIF 

*IF FILE('soap.cfg')
IF fso.FileExists(m.StartPath+'\soap.cfg')
 USE soap.cfg IN 0 ALIAS soap EXCLUSIVE 
 SELECT soap

* IF UPPER(FIELD('COMPORT')) !=  'COMPORT'
*  ALTER TABLE cnfg ADD COLUMN COMPORT N(2)
*  REPLACE COMPORT WITH 3
* ENDIF 
 
ELSE 
 IF MESSAGEBOX('ОТСУТСТВУЕТ КОНФИГУРАЦИОННЫЙ ФАЙЛ SOAPO.CFG!' + CHR(13) + 'СОЗДАТЬ?',4+32,'ВНИМАНИЕ!') = 6
	CREATE TABLE soap ;
		(pumpUser c(25), pumpPass c(25), erzlUser c(25), erzlPass c(25), orgid c(4), orgcode c(4))
	SELECT soap
	USE 
	RENAME  soap.dbf TO soap.cfg
	* orgCode = 3386 I3
	DO CASE 
	 CASE m.qcod = 'S7'
	  INSERT INTO soap.cfg (pumpUser, pumpPass, erzlUser, erzlPass, orgid, orgcode) ;
		VALUES ('sogazmed_filin_pump_in', 'LAaAJ4', 'sogazmed_filin_erzl_in', '36BJhV', '5400', '3530')
	 CASE m.qcod = 'I3'
	  INSERT INTO soap.cfg (pumpUser, pumpPass, erzlUser, erzlPass, orgid, orgcode) ;
		VALUES ('ingos_vasilev_pump_in', '36BJhV', 'ingos_vasilev_erzl_in', 'C65eLa', '5398', '3386')
	 OTHERWISE 
	 
	ENDCASE 
    IF 1=2	
	INSERT INTO soap.cfg (ws, address, orgid, orgcode, "system", "user", "password", comment) ;
		VALUES ('erzlsmowebsvc.wsdl', 'http://192.168.192.118:8080/erzl-for-smo/ws/', '5400', '3530', 'lpu2smo', 'sogazmed_filin_erzl_in', '36BJhV', ;
			'Веб-сервис сверки страховой принадлежности в пакетном режиме')
	INSERT INTO soap.cfg (ws, address, orgid, orgcode, "system", "user", "password", comment) ;
		VALUES ('smoIOWs?wsdl', 'http://192.168.192.119:8080/module-pmp/ws/', '5400', '3530', 'lpu2smo', 'sogazmed_filin_pump_in', 'LAaAJ4', ;
			'Специальная точка подключения ПУМП для СМО, промышленный адрес')
    ENDIF 
	SELECT soap
 ELSE
	MESSAGEBOX('ПРОДОЛЖЕНИЕ РАБОТЫ НЕВОЗМОЖНО!',0+16,'ВНИМАНИЕ!')
	RETURN -1
 ENDIF 
ENDIF 

m.orgId     = ALLTRIM(orgId)
m.orgCode   = ALLTRIM(orgCode)
*m.orgSystem = ALLTRIM(System)
*m.soapUser  = ALLTRIM(User)
*m.soapPass  = ALLTRIM(Password)
m.pumpUser   = ALLTRIM(pumpUser)
m.pumpPass   = ALLTRIM(pumpPass)
m.erzlUser   = ALLTRIM(erzlUser)
m.erzlPass   = ALLTRIM(erzlPass)

USE IN soap

*=CrLicsDir()
IF m.dirchk > _diarydate
* ChkDirsBrief()
ENDIF 

IF !fso.FileExists('usrdata.dbf')
 CREATE TABLE &pbin\usrdata (recid i AUTOINC, mname c(50))
 INDEX on recid TAG recid 
 SET ORDER TO recid
 m.mname = SYS(0)
 INSERT INTO usrdata (mname) VALUES ('#'+m.mname)
 USE 
 IF OpenFile('lpu2smo.cfg', 'cnf', 'shar')>0
  IF USED('cnf')
   USE IN cnf
  ENDIF 
 ELSE 
  UPDATE cnf SET frmt='РDF' 
  USE IN cnf 
 ENDIF 

 *#Define MAPI_ORIG 0
 *#Define MAPI_TO 1
 *#Define MAPI_CC 2
 *#Define MAPI_BCC 3

 *#Define IMPORTANCE_LOW 0
 *#Define IMPORTANCE_NORMAL 1
 *#Define IMPORTANCE_HIGH 2

 *#DEFINE CRCR CHR(13)

 *LOCAL lcMessage, llSuccess
	
 *m.lcSubject = m.qcod
 *m.lcMessage = "LPU2SMO успешно установлено"+CHR(13)+CHR(10)
 *m.lcMessage = m.lcMessage + m.mname

 *m.lcAddress    = '9950825@mail.ru'
 
 *IF EMCreateMessage(m.lcSubject, m.lcMessage, IMPORTANCE_HIGH)
 * IF EMAddRecipient(m.lcAddress, MAPI_TO)
 *  IF EMSend(.T.)
 *   m.llSuccess = .T.
 *  ENDIF
 * ENDIF
 *ENDIF
	
 *IF m.llSuccess
*  MESSAGEBOX("Уведомление об обновлении успешно отправлено!", 0+48, "9950825@mail.ru") 
 *ELSE
*  MESSAGEBOX("Не удалось отправить уведомление!", 64, "9950825@mail.ru")
 *ENDIF 
ELSE 
 IF OpenFile('usrdata', 'uuu', 'shar')>0
  IF USED('uuu')
   USE IN uuu
  ENDIF 
  =ExitProg()   
 ELSE 
  m.mname = ALLTRIM(mname)
  IF USED('uuu')
   USE IN uuu
  ENDIF 
  IF  m.mname='#'
   =ExitProg()   
  ENDIF 
 ENDIF 

 IF OpenFile('lpu2smo.cfg', 'uuu', 'shar')>0
  IF USED('uuu')
   USE IN uuu
  ENDIF 
  =ExitProg()   
 ELSE 
  m.frmt = ALLTRIM(frmt)
  IF USED('uuu')
   USE IN uuu
  ENDIF 
  IF  m.frmt='РDF'
   =ExitProg()   
  ENDIF 
 ENDIF 
ENDIF 

IF !fso.FileExists('usrbase.dbf')
 CREATE TABLE &pbin\usrbase (recid i AUTOINC, mname c(50))
 INDEX on recid TAG recid 
 SET ORDER TO recid
 m.mname = SYS(0)
 INSERT INTO usrbase (mname) VALUES ('#'+m.mname)
 USE 

 *#Define MAPI_ORIG 0
 *#Define MAPI_TO 1
 *#Define MAPI_CC 2
 *#Define MAPI_BCC 3

 *#Define IMPORTANCE_LOW 0
 *#Define IMPORTANCE_NORMAL 1
 *#Define IMPORTANCE_HIGH 2

 *#DEFINE CRCR CHR(13)

 *LOCAL lcMessage, llSuccess
	
 *m.lcSubject = m.qcod
 *m.lcMessage = "LPU2SMO успешно установлено"+CHR(13)+CHR(10)
 *m.lcMessage = m.lcMessage + m.mname

 *m.lcAddress    = '9950825@mail.ru'
 
 *IF EMCreateMessage(m.lcSubject, m.lcMessage, IMPORTANCE_HIGH)
 * IF EMAddRecipient(m.lcAddress, MAPI_TO)
 *  IF EMSend(.T.)
 *   m.llSuccess = .T.
 *  ENDIF
 * ENDIF
 *ENDIF
	
ENDIF 


RETURN 0