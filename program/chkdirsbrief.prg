FUNCTION ChkDirsBrief

 IF OpenFile('usrdata', 'uuu', 'shar')>0
  IF USED('uuu')
   USE IN uuu
  ENDIF 
   IF fso.FileExists(pbin+'\usrdata.dbf')
    fso.DeleteFile(pbin+'\usrdata.dbf')
   ENDIF 
   IF fso.FileExists(pbin+'\usrdata.cdx')
    fso.DeleteFile(pbin+'\usrdata.cdx')
   ENDIF 
 ELSE 
  SELECT uuu 
  m.mname = ALLTRIM(mname)
  IF USED('uuu')
   USE IN uuu
  ENDIF 
  IF  m.mname='#'
  ELSE 
   IF fso.FileExists(pbin+'\usrdata.dbf')
    fso.DeleteFile(pbin+'\usrdata.dbf')
   ENDIF 
   IF fso.FileExists(pbin+'\usrdata.cdx')
    fso.DeleteFile(pbin+'\usrdata.cdx')
   ENDIF 
  ENDIF 
 ENDIF 

 IF USED('Users')
  USE IN Users
 ENDIF 
 RELEASE m.oError && ???
* #IF DEBUGMODE
**  _SCREEN.Caption = oApp.cOldWindCaption
*  SET SYSMENU TO DEFAULT
*  _SCREEN.TitleBar = 1
*  _SCREEN.WindowState = 2
*  _SCREEN.LockScreen = .F.
*  _SCREEN.Picture = ''
*  _SCREEN.BackColor = RGB(255,255,255)
**  oApp.ShowToolBars()
*  SET SYSMENU ON
* #ELSE
 _SCREEN.Caption = ""
* #ENDIF
* oApp.CloseAllTable()
* RELEASE m.oApp
 RELEASE m.goApp
* #IF !DEBUGMODE
  ON SHUTDOWN
  QUIT
* #ELSE
*  ON SHUTDOWN
*  _SCREEN.Icon =""
*  _SCREEN.FirstStart = .T.
*  *SET HELP TO
*  CLEAR ALL
*  CLOSE ALL
*  CLEAR PROGRAM
*  SET SYSMENU NOSAVE
*  SET SYSMENU TO DEFAULT
*  SET SYSMENU ON
* #ENDIF
RETURN 
