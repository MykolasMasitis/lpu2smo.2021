PROCEDURE m_menu
SET SYSMENU TO

DEFINE PAD mmenu_1  OF _MSYSMENU PROMPT '\<хмтнплюжхъ нр кос' COLOR SCHEME 3 KEY ALT+A, ""
DEFINE PAD mmenu_3  OF _MSYSMENU PROMPT '\<щйяоепрхгю' COLOR SCHEME 3 KEY ALT+C , ""
DEFINE PAD mmenu_4  OF _MSYSMENU PROMPT '\<яопюбнвмхйх' COLOR SCHEME 3 KEY ALT+C , ""
DEFINE PAD mmenu_6  OF _MSYSMENU PROMPT '\<тхмюмяш' COLOR SCHEME 3 KEY ALT+E , ""
DEFINE PAD mmenu_7  OF _MSYSMENU PROMPT '\<яепбхя' COLOR SCHEME 3 KEY ALT+F , ""
DEFINE PAD mmenu_8  OF _MSYSMENU PROMPT '\<яепбхя-2' COLOR SCHEME 3 KEY ALT+F , ""
DEFINE PAD mmenu_9  OF _MSYSMENU PROMPT '\<POSTGRESQL' COLOR SCHEME 3 KEY ALT+F , ""
ON PAD mmenu_1  OF _MSYSMENU ACTIVATE POPUP popInfFrLpu
ON PAD mmenu_3  OF _MSYSMENU ACTIVATE POPUP popMEE
ON PAD mmenu_4  OF _MSYSMENU ACTIVATE POPUP popMedSpr
ON PAD mmenu_6  OF _MSYSMENU ACTIVATE POPUP popBuch
ON PAD mmenu_7  OF _MSYSMENU ACTIVATE POPUP popTuneUp
ON PAD mmenu_8  OF _MSYSMENU ACTIVATE POPUP popTuneUp2
ON PAD mmenu_9  OF _MSYSMENU ACTIVATE POPUP popPostgreSQL

DEFINE POPUP popInfFrLpu MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popInfFrLpu PROMPT 'ядекюрэ оепед гюйпшрхел оепхндю'
DEFINE BAR 02 OF popInfFrLpu PROMPT 'ядекюрэ оепед нйрпшрхел оепхндю'
DEFINE BAR 03 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 04 OF popInfFrLpu PROMPT 'бхрпхмю елхюя'
DEFINE BAR 05 OF popInfFrLpu PROMPT 'бхрпхмю SOAP'
DEFINE BAR 06 OF popInfFrLpu PROMPT 'фспмюк гюопнянб й епгк (SOAP)'
DEFINE BAR 07 OF popInfFrLpu PROMPT 'тнплхпнбюмхе UD-тюикнб'
DEFINE BAR 08 OF popInfFrLpu PROMPT 'тнплхпнбюмхе UP-тюикнб'
DEFINE BAR 09 OF popInfFrLpu PROMPT 'яопюбнвмхй днцнбнпнб лн'
DEFINE BAR 10 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 11 OF popInfFrLpu PROMPT 'ябндмши явер гю оепхнд'
DEFINE BAR 12 OF popInfFrLpu PROMPT 'ябндмши явер гю цнд' SKIP FOR m.IsNotePad
DEFINE BAR 13 OF popInfFrLpu PROMPT 'янапюрэ ябндмши тюик цняохрюкхгюжхи'
DEFINE BAR 14 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 15 OF popInfFrLpu PROMPT 'ятнплхпнбюрэ CTRL-тюикш' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 16 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 17 OF popInfFrLpu PROMPT 'дхяоюмяепхгюжхъ'
DEFINE BAR 18 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 19 OF popInfFrLpu PROMPT 'оепянрвер' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 20 OF popInfFrLpu PROMPT 'тхмтюик' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 21 OF popInfFrLpu PROMPT 'ME-тюикш' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 22 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 23 OF popInfFrLpu PROMPT 'явер мю оюжхемрю'
DEFINE BAR 24 OF popInfFrLpu PROMPT 'яопюбнвмхй онкэгнбюрекеи'
DEFINE BAR 25 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 26 OF popInfFrLpu PROMPT 'оюйермюъ оевюрэ'
DEFINE BAR 27 OF popInfFrLpu PROMPT 'оепемня бяеу днйслемрнб б ндмс оюойс'
DEFINE BAR 28 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 29 OF popInfFrLpu PROMPT 'бшунд'

ON BAR 01 OF popInfFrLpu ACTIVATE POPUP DoBeforeClosePeriod
ON BAR 02 OF popInfFrLpu ACTIVATE POPUP DoBeforeOpenPeriod
ON SELECTION BAR 04 OF popInfFrLpu DO FORM MailView
ON SELECTION BAR 05 OF popInfFrLpu DO FORM MailSoap
ON SELECTION BAR 06 OF popInfFrLpu DO FORM ErzSoap
ON SELECTION BAR 07 OF popInfFrLpu do MakeUDFilesN
ON SELECTION BAR 08 OF popInfFrLpu do MakeUPFilesN
ON SELECTION BAR 09 OF popInfFrLpu DO FORM ViewDogs
ON SELECTION BAR 11 OF popInfFrLpu DO FORM ViewPeriod
ON SELECTION BAR 12 OF popInfFrLpu DO FORM ViewSvYear
ON SELECTION BAR 13 OF popInfFrLpu DO MakeSvGsp
ON SELECTION BAR 15 OF popInfFrLpu DO MakeCtrls
ON BAR 17 OF popInfFrLpu ACTIVATE POPUP DispMenu
ON BAR 19 OF popInfFrLpu ACTIVATE POPUP  m_pers
ON BAR 20 OF popInfFrLpu ACTIVATE POPUP  m_finfile
ON BAR 21 OF popInfFrLpu ACTIVATE POPUP  m_mefiles
ON SELECTION BAR 23 OF popInfFrLpu DO FindPaz
ON SELECTION BAR 24 OF popInfFrLpu DO form sprusers
ON BAR 26 OF popInfFrLpu ACTIVATE POPUP popPrn
ON SELECTION BAR 27 OF popInfFrLpu do CopyAllDocs
ON SELECTION BAR 29 OF popInfFrLpu clea events 

DEFINE POPUP DoBeforeClosePeriod MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF DoBeforeClosePeriod PROMPT 'ятнплхпнбюрэ тюик DSP (dsp.dbf)'
DEFINE BAR 2 OF DoBeforeClosePeriod PROMPT 'янапюрэ тюик "жемрп гднпнбэъ" (nsi\polic_h.dbf)'
DEFINE BAR 3 OF DoBeforeClosePeriod PROMPT 'янапюрэ D-тюик (deads.dbf)'
DEFINE BAR 4 OF DoBeforeClosePeriod PROMPT '\-'
DEFINE BAR 5 OF DoBeforeClosePeriod PROMPT 'юйрсюкхгхпнбюрэ мях'
ON SELECTION BAR 1 OF DoBeforeClosePeriod DO MakeDspFile with .t., .f.
ON SELECTION BAR 2 OF DoBeforeClosePeriod DO Make15001
ON SELECTION BAR 3 OF DoBeforeClosePeriod DO MakeDeads
ON SELECTION BAR 5 OF DoBeforeClosePeriod DO ActNSI

DEFINE POPUP DoBeforeOpenPeriod MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF DoBeforeOpenPeriod PROMPT 'йнмбепрхпнбюрэ тюик medicament.xml'
DEFINE BAR 2 OF DoBeforeOpenPeriod PROMPT 'йнмбепрхпнбюрэ тюик medicament_man_pack.xml'
DEFINE BAR 3 OF DoBeforeOpenPeriod PROMPT 'йнмбепрхпнбюрэ тюик medicament_mfc.xml'
DEFINE BAR 4 OF DoBeforeOpenPeriod PROMPT 'йнмбепрхпнбюрэ тюик F003.xml'
DEFINE BAR 5 OF DoBeforeOpenPeriod PROMPT '\-'
DEFINE BAR 6 OF DoBeforeOpenPeriod PROMPT 'гюцпсгхрэ тнплс нля-15/3Ю (репюохъ)'
DEFINE BAR 7 OF DoBeforeOpenPeriod PROMPT 'гюцпсгхрэ тнплс нля-15/3Ю (ярнлюрнкнцхъ)'
DEFINE BAR 8 OF DoBeforeOpenPeriod PROMPT 'гюцпсгхрэ тнплс нля-15/3Х (Я 01.01.2020)'
DEFINE BAR 9 OF DoBeforeOpenPeriod PROMPT '\-'
DEFINE BAR 10 OF DoBeforeOpenPeriod PROMPT 'гюцпсгхрэ мнлепмхй'
DEFINE BAR 11 OF DoBeforeOpenPeriod PROMPT 'гюцпсгхрэ ярно-кхяр'
ON SELECTION BAR 1 OF DoBeforeOpenPeriod DO Medicament_sax
ON SELECTION BAR 2 OF DoBeforeOpenPeriod DO MedPack_sax
ON SELECTION BAR 3 OF DoBeforeOpenPeriod DO MedMFC_sax
ON SELECTION BAR 4 OF DoBeforeOpenPeriod DO F003_sax
ON SELECTION BAR 6 OF DoBeforeOpenPeriod DO AppOMS15
ON SELECTION BAR 7 OF DoBeforeOpenPeriod DO AppOMS15st
ON SELECTION BAR 8 OF DoBeforeOpenPeriod DO AppOMS15i
ON SELECTION BAR 10 OF DoBeforeOpenPeriod DO MakeOutS
ON SELECTION BAR 11 OF DoBeforeOpenPeriod DO LoadStopList

DEFINE POPUP m_pers MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF m_pers PROMPT 'ятнплхпнбюрэ оепянрвер'
DEFINE BAR 2 OF m_pers PROMPT 'нропюбхрэ оепянрвер'
DEFINE BAR 3 OF m_pers PROMPT 'онхяй нрбернб лцтнля мю оепянрвер' SKIP FOR !INLIST(gcUser,'OMS','USR')
ON SELECTION BAR 1 OF m_pers do MakeYFiles
ON SELECTION BAR 2 OF m_pers do SendPers
ON SELECTION BAR 3 OF m_pers DO FindPersAns

DEFINE POPUP m_finfile MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF m_finfile PROMPT 'ятнплхпнбюрэ тхмтюик'
DEFINE BAR 2 OF m_finfile PROMPT 'нропюбхрэ тхмтюик'
DEFINE BAR 3 OF m_finfile PROMPT 'онхяй нрберю лцтнля мю тхмтюик' SKIP
ON SELECTION BAR 1 OF m_finfile DO IIF(INLIST(m.qcod,'S7','R2'), 'MakeFinS7', 'MakeFin')
ON SELECTION BAR 2 OF m_finfile do SendFinFile

DEFINE POPUP m_mefiles MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF m_mefiles PROMPT 'ятнплхпнбюрэ ле-тюикш'
DEFINE BAR 2 OF m_mefiles PROMPT 'нропюбхрэ ле-тюикш'
DEFINE BAR 3 OF m_mefiles PROMPT 'онхяй нрбернб лцтнля мю ле-тюикш' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 4 OF m_mefiles PROMPT 'онхяй CTRL мю ле-тюикш' SKIP FOR !INLIST(gcUser,'OMS','USR')
ON SELECTION BAR 1 OF m_mefiles do IIF(m.qcod='I3', 'MakeMEFilesI3', 'MakeMEFiles')
ON SELECTION BAR 2 OF m_mefiles do SendMEFiles
ON SELECTION BAR 3 OF m_mefiles DO FindMeAns
ON SELECTION BAR 4 OF m_mefiles DO FindMeCtrls

DEFINE POPUP popWebServices MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF popWebServices PROMPT 'getBillStatuses (БЯЕ лн)'
DEFINE BAR 2 OF popWebServices PROMPT 'getMailGw (БЯЕ лн)'
ON SELECTION BAR 1 OF popWebServices getBillStatuses(0, null, .f., 'SMO')
ON SELECTION BAR 2 OF popWebServices getMailGw(0, null, "")

DEFINE POPUP popPrn MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF popPrn PROMPT 'оевюрэ опнрнйнкнб опх╗лйх яверю'
DEFINE BAR 2 OF popPrn PROMPT 'оевюрэ юйрнб лщй'
DEFINE BAR 3 OF popPrn PROMPT 'оевюрэ юйрнб на нокюре он ондсьебнцн тхмюмяхпнбюмхъ'
ON SELECTION BAR 1 OF popPrn DO PackPrnPr
ON SELECTION BAR 2 OF popPrn DO PackPrnMc
ON SELECTION BAR 3 OF popPrn DO PackPrnPdf

DEFINE POPUP DispMenu MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF DispMenu PROMPT 'тнплхпнбюмхе тюикю DSP'
DEFINE BAR 2 OF DispMenu PROMPT '\-'
DEFINE BAR 3 OF DispMenu PROMPT 'тнплхпнбюмхе нрверю'
DEFINE BAR 4 OF DispMenu PROMPT 'тнплхпнбюмхе нрверю (оПНТн_б_032018_50046)'
DEFINE BAR 5 OF DispMenu PROMPT 'тнплхпнбюмхе нрверю (дб_1_032018_50046)'
DEFINE BAR 6 OF DispMenu PROMPT 'тнплхпнбюмхе нрверю (дб_1_032018_50046, бепяхъ 2)'
DEFINE BAR 7 OF DispMenu PROMPT 'тнплхпнбюмхе нрверю (опхкнфемхе 4)'
DEFINE BAR 8 OF DispMenu PROMPT '\-'
DEFINE BAR 9 OF DispMenu PROMPT 'тнплхпнбюмхе нрвернб б кос'
DEFINE BAR 10 OF DispMenu PROMPT '\-'
DEFINE BAR 11 OF DispMenu PROMPT 'йнппейрхпнбйю йнднб кос'
DEFINE BAR 12 OF DispMenu PROMPT '\-'
DEFINE BAR 13 OF DispMenu PROMPT 'йнппейрхпнбйю йнднб кос-2'
DEFINE BAR 14 OF DispMenu PROMPT '\-'
DEFINE BAR 15 OF DispMenu PROMPT 'нрвермнярэ опнт'
DEFINE BAR 16 OF DispMenu PROMPT '\-'
DEFINE BAR 17 OF DispMenu PROMPT 'янапюрэ тюик "жемрп гднпнбэъ"'
DEFINE BAR 18 OF DispMenu PROMPT 'янапюрэ пееярп дхяоюмяепмшу анкэмшу'
DEFINE BAR 19 OF DispMenu PROMPT 'дм_рЕПЮОХЪ'
DEFINE BAR 20 OF DispMenu PROMPT 'дм_нМЙНКНЦХЪ'

ON SELECTION BAR 1 OF DispMenu DO MakeDspFile with .t., .f.
ON SELECTION BAR 3 OF DispMenu DO NewDspMonitorN
ON SELECTION BAR 4 OF DispMenu DO DspMonitorProf
ON SELECTION BAR 5 OF DispMenu DO DspMonitorDV
ON SELECTION BAR 6 OF DispMenu DO DspMonitorDV2
ON SELECTION BAR 7 OF DispMenu DO DspMonitorXX
ON SELECTION BAR 9 OF DispMenu DO FormDDDS
ON SELECTION BAR 11 OF DispMenu DO CorrDsp
ON SELECTION BAR 13 OF DispMenu DO CorrMcod
ON BAR 15 OF DispMenu ACTIVATE POPUP ProfReps
ON SELECTION BAR 17 OF DispMenu DO Make15001
ON SELECTION BAR 18 OF DispMenu DO MakeDispReestr
ON SELECTION BAR 19 OF DispMenu DO FormDNTherapy
ON SELECTION BAR 20 OF DispMenu DO FormDNOncology

DEFINE POPUP ProfReps MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF ProfReps PROMPT 'рюакхжю 1'
DEFINE BAR 2 OF ProfReps PROMPT 'рюакхжю 2'
DEFINE BAR 3 OF ProfReps PROMPT 'рюакхжю 3'
ON SELECTION BAR 1 OF ProfReps DO ProfRepT1
ON SELECTION BAR 2 OF ProfReps DO ProfRepT2
ON SELECTION BAR 3 OF ProfReps DO ProfRepT3

DEFINE POPUP AccsPeriod MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF AccsPeriod PROMPT 'ябндмши явер'
ON SELECTION BAR 1 OF AccsPeriod DO FORM ViewPeriod

DEFINE POPUP AccsYear MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF AccsYear PROMPT 'пецхярп'
DEFINE BAR 2 OF AccsYear PROMPT 'сяксцх'
DEFINE BAR 3 OF AccsYear PROMPT 'йнлахмюжхъ'
ON SELECTION BAR 3 OF AccsYear DO FORM ViewSvYear

DEFINE POPUP popInfToMGF MARGIN RELATIVE SHADOW COLOR SCHEME 4
*DEFINE BAR 01 OF popInfToMGF PROMPT 'юцпецхпнбюммше яверю' PICTURE 'GROUP3.BMP' SKIP 
*DEFINE BAR 02 OF popInfToMGF PROMPT 'ятнплхпнбюрэ тхм-тюик' PICTURE 'GROUP3.BMP'
*DEFINE BAR 03 OF popInfToMGF PROMPT 'оепянмхтхжхпнбюммше яверю' PICTURE 'GROUP3.BMP' SKIP 
*DEFINE BAR 04 OF popInfToMGF PROMPT '\-'
*DEFINE BAR 01 OF popInfToMGF PROMPT 'тнплю ╧1'
*DEFINE BAR 06 OF popInfToMGF PROMPT 'тнплю "оц" (лщй)'
*DEFINE BAR 07 OF popInfToMGF PROMPT '\-' 
*DEFINE BAR 08 OF popInfToMGF PROMPT 'тнплю "оц" (лщщ ОКЮМНБЮЪ)'
*DEFINE BAR 09 OF popInfToMGF PROMPT 'тнплю "оц" (лщщ ЖЕКЕБЮЪ)'
*DEFINE BAR 10 OF popInfToMGF PROMPT 'тнплю "оц" (лщщ РЕЛЮРХВЕЯЙЮЪ)'
*DEFINE BAR 11 OF popInfToMGF PROMPT 'тнплю "оц" (лщщ ОН ФЮКНАЮЛ)'
*DEFINE BAR 12 OF popInfToMGF PROMPT '\-' 
*DEFINE BAR 13 OF popInfToMGF PROMPT 'тнплю "оц" (щйло ОКЮМНБЮЪ)'
*DEFINE BAR 14 OF popInfToMGF PROMPT 'тнплю "оц" (щйло ЖЕКЕБЮЪ)'
*DEFINE BAR 15 OF popInfToMGF PROMPT 'тнплю "оц" (щйло РЕЛЮРХВЕЯЙЮЪ)'
*DEFINE BAR 16 OF popInfToMGF PROMPT 'тнплю "оц" (щйло ОН ФЮКНАЮЛ)'
*DEFINE BAR 17 OF popInfToMGF PROMPT '\-' 
*DEFINE BAR 01 OF popInfToMGF PROMPT 'напюыюелнярэ гюярпюунбюммшу' 
*DEFINE BAR 02 OF popInfToMGF PROMPT '\-' 
*DEFINE BAR 03 OF popInfToMGF PROMPT 'ябепхрэ я мнлепмхйнл' 
*DEFINE BAR 04 OF popInfToMGF PROMPT '\-' 
*DEFINE BAR 01 OF popInfToMGF PROMPT 'нрвер он онкс/бнгпюярс (демэцх)' 
*DEFINE BAR 02 OF popInfToMGF PROMPT 'нрвер он онкс/бнгпюярс (кчдх)' 
*DEFINE BAR 03 OF popInfToMGF PROMPT 'нрвер он онкс/бнгпюярс (демэцх) ярюж' 
*DEFINE BAR 04 OF popInfToMGF PROMPT 'нрвер он онкс/бнгпюярс (кчдх) ярюж' 

*ON SELECTION BAR 01 OF popInfToMGF goApp.doForm('frm_agreg','mgfoms')
*ON SELECTION BAR 02 OF popInfToMGF DO MakeFin
*ON SELECTION BAR 03 OF popInfToMGF DO MakeYFiles
*ON SELECTION BAR 01 OF popInfToMGF DO FormN1
*ON SELECTION BAR 06 OF popInfToMGF DO FormPGMek
*ON SELECTION BAR 08 OF popInfToMGF DO FormPGMee WITH 2
*ON SELECTION BAR 09 OF popInfToMGF DO FormPGMee WITH 3
*ON SELECTION BAR 10 OF popInfToMGF DO FormPGMee WITH 7
*ON SELECTION BAR 11 OF popInfToMGF DO FormPGMee WITH 8
*ON SELECTION BAR 13 OF popInfToMGF DO FormPGMee WITH 4
*ON SELECTION BAR 14 OF popInfToMGF DO FormPGMee WITH 5
*ON SELECTION BAR 15 OF popInfToMGF DO FormPGMee WITH 6
*ON SELECTION BAR 16 OF popInfToMGF DO FormPGMee WITH 9
*ON SELECTION BAR 01 OF popInfToMGF DO ObrPrikl
*ON SELECTION BAR 03 OF popInfToMGF DO SvOutS
*ON SELECTION BAR 01 OF popInfToMGF DO SagOpl
*ON SELECTION BAR 02 OF popInfToMGF DO SagOpl2
*ON SELECTION BAR 03 OF popInfToMGF DO SagOpls
*ON SELECTION BAR 04 OF popInfToMGF DO SagOpl2s

DEFINE POPUP popMEE MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popMEE PROMPT 'щйяоепрхгю рейсыецн оепхндю'
DEFINE BAR 02 OF popMEE PROMPT 'щйяоепрхгю опнхгбнкэмнцн оепхндю'
DEFINE BAR 03 OF popMEE PROMPT 'йпхрепхх нранпю дкъ щйяоепрхгш'
DEFINE BAR 04 OF popMEE PROMPT 'йпхрепхх нранпю дкъ щйяоепрхгш (NEW)'
DEFINE BAR 05 OF popMEE PROMPT '\-'
DEFINE BAR 06 OF popMEE PROMPT 'тнплхпнбюмхе ME-тюикнб'
DEFINE BAR 07 OF popMEE PROMPT 'фспмюк ябндмшу юйрнб'
DEFINE BAR 08 OF popMEE PROMPT 'фспмюк юйрнб ярпюунбшу яксвюеб'
DEFINE BAR 09 OF popMEE PROMPT '\-'
DEFINE BAR 10 OF popMEE PROMPT 'фспмюк пееярпнб юйрнб'
DEFINE BAR 11 OF popMEE PROMPT 'фспмюк гюопнянб олд'
DEFINE BAR 12 OF popMEE PROMPT '\-'
DEFINE BAR 13 OF popMEE PROMPT 'ятнплхпнбюрэ тюикш хлонпрю' PICTURE 'GROUP3.BMP'
DEFINE BAR 14 OF popMEE PROMPT 'хлонпрхпнбюрэ юйрш щйяоепрхгш' PICTURE 'GROUP3.BMP'
DEFINE BAR 15 OF popMEE PROMPT 'яоежхлонпр (хмцняярпюу-л)' PICTURE 'GROUP3.BMP' SKIP FOR m.qcod!='I3'
DEFINE BAR 16 OF popMEE PROMPT '\-'
DEFINE BAR 17 OF popMEE PROMPT 'гюцпсгхрэ тюикш хлонпрю' PICTURE 'GROUP2.BMP'
DEFINE BAR 18 OF popMEE PROMPT 'щйяонпрхпнбюрэ юйрш щйяоепрхгш' PICTURE 'GROUP2.BMP'
DEFINE BAR 19 OF popMEE PROMPT 'яоежщйяонпр (йпюрйхи)' PICTURE 'GROUP2.BMP' SKIP FOR m.qcod!='I3'
DEFINE BAR 20 OF popMEE PROMPT 'яоежщйяонпр (онкмши)' PICTURE 'GROUP2.BMP' SKIP FOR m.qcod!='I3'
DEFINE BAR 21 OF popMEE PROMPT '\-'
DEFINE BAR 22 OF popMEE PROMPT 'опнбепйю яннрберярбхъ яслл ямърхи'
DEFINE BAR 23 OF popMEE PROMPT '\-'
DEFINE BAR 24 OF popMEE PROMPT 'ледхйн-щйнмнлхвеяйхе тнплш S7'
DEFINE BAR 25 OF popMEE PROMPT 'ледхйн-щйнмнлхвеяйхе тнплш S2'
DEFINE BAR 26 OF popMEE PROMPT 'гюопня ттнля нр 18.02.2016'
DEFINE BAR 27 OF popMEE PROMPT 'гюопня ттнля нр 18.02.2016 (лщй)'
DEFINE BAR 28 OF popMEE PROMPT '\-'
DEFINE BAR 29 OF popMEE PROMPT 'япюбмемхе пегскэрюрнб гюопнянб'
DEFINE BAR 30 OF popMEE PROMPT 'опхкнфемхе ╧2 (лнмхрнпхмц, нмйнкнцхъ)'
DEFINE BAR 31 OF popMEE PROMPT 'опхкнфемхе ╧2 (лнмхрнпхмц, нмйнкнцхъ, бепяхъ хмцняя)'
DEFINE BAR 32 OF popMEE PROMPT 'то "аНПЭАЮ Я НМЙНКНЦХВЕЯЙХЛХ ГЮАНКЕБЮМХЪЛХ"'
DEFINE BAR 33 OF popMEE PROMPT '\-'
DEFINE BAR 34 OF popMEE PROMPT 'нрвер он ямърхъл PPA'
DEFINE BAR 35 OF popMEE PROMPT 'нрвер дкъ хбюмю тхкхмю'

ON SELECTION BAR 01 OF PopMEE DO viewmee
ON SELECTION BAR 02 OF PopMEE DO FORM ViewYear
ON BAR 03 OF PopMEE ACTIVATE POPUP ExpCriteria
ON BAR 04 OF PopMEE ACTIVATE POPUP ExpCritNew
ON SELECTION BAR 06 OF popMEE DO IIF(m.qcod='I3', 'MakeMEFilesI3', 'MakeMEFiles')
ON SELECTION BAR 07 OF popMEE DO FORM ViewActSV
ON SELECTION BAR 08 OF popMEE DO FORM ViewActSS
ON SELECTION BAR 10 OF popMEE DO FORM ViewRss
ON SELECTION BAR 11 OF popMEE DO FORM ViewRqst
ON SELECTION BAR 13 OF popMEE DO ImpExp
ON SELECTION BAR 14 OF popMEE DO ImpActs
ON SELECTION BAR 15 OF popMEE DO ImpExpI3
ON SELECTION BAR 17 OF popMEE DO ExpExp
ON SELECTION BAR 18 OF popMEE DO ExpActs
ON SELECTION BAR 19 OF popMEE DO ExpExpI3c
ON SELECTION BAR 20 OF popMEE DO ExpExpI3
ON SELECTION BAR 22 OF popMEE DO CmpMee
ON BAR 24 OF PopMEE ACTIVATE POPUP popS7
ON BAR 25 OF PopMEE ACTIVATE POPUP popS2
ON SELECTION BAR 26 OF popMEE DO FFOMS18022016
ON SELECTION BAR 27 OF popMEE DO FFOMS18022016mek
ON SELECTION BAR 29 OF popMEE DO CompRequests
ON SELECTION BAR 30 OF popMEE DO Onk_pril2
ON SELECTION BAR 31 OF popMEE DO Onk_pril2_v2
ON SELECTION BAR 32 OF popMEE DO FPOnko
ON SELECTION BAR 34 OF popMEE DO RepPPA
ON SELECTION BAR 35 OF popMEE DO RepIVA001

DEFINE POPUP popS7 MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popS7 PROMPT 'тнплю ь-0 (нрвер дкъ йпс)'
DEFINE BAR 02 OF popS7 PROMPT 'тнплю ь-1' SKIP && АПЕД!
DEFINE BAR 03 OF popS7 PROMPT 'тнплю ь-2 (нрвер он пндднлюл)' SKIP && Я 2014 МХЙРН МЕ ОНКЭГНБЮКЯЪ!
DEFINE BAR 04 OF popS7 PROMPT 'тнплю ь-3 (нрвер он опнбедеммшл щйяоепрхгюл)' SKIP  && БПЪД КХ ХЯОНКЭГСЕРЯЪ
DEFINE BAR 05 OF popS7 PROMPT 'тнплю ь-3ахя (нрвер он опнбедеммшл щйяоепрхгюл)' SKIP && БПЪД КХ ХЯОНКЭГСЕРЯЪ
DEFINE BAR 06 OF popS7 PROMPT 'тнплю ь-4 лщщ (опхкнфемхе 8)' SKIP && Я ЪМБЮПЪ 2014 МХЙРН МЕ ОНКЭГНБЮКЯЪ
DEFINE BAR 07 OF popS7 PROMPT 'тнплю ь-4 щйло (опхкнфемхе 8)' SKIP && Я ЪМБЮПЪ 2014 МХЙРН МЕ ОНКЭГНБЮКЯЪ
DEFINE BAR 08 OF popS7 PROMPT 'тнплю ь-5 (ярюрхярхйю ямърхи он лщй)' && МСФМЮ!
DEFINE BAR 09 OF popS7 PROMPT 'тнплю  ю3-0' && МСФМЮ!
DEFINE BAR 10 OF popS7 PROMPT 'ярюрхярхйю ямърхи б пюгпеге лщй,лщщ,щйло'
DEFINE BAR 11 OF popS7 PROMPT 'пюяьхпеммюъ ярюрхярхйю ямърхи б пюгпеге лщй,лщщ,щйло'
DEFINE BAR 12 OF popS7 PROMPT '\-'
DEFINE BAR 13 OF popS7 PROMPT 'тнплю ь-6 (пюгдек II тнплш 1)'
DEFINE BAR 14 OF popS7 PROMPT 'тнплю ь-7 (нрвер он жекебни щйяоепрхге)'
DEFINE BAR 15 OF popS7 PROMPT 'тнплю ь-8'
DEFINE BAR 16 OF popS7 PROMPT 'нярпши йнпнмюпмши яхмдпнл'
DEFINE BAR 17 OF popS7 PROMPT 'нярпне мюпсьемхе лнгцнбнцн йпнбннапюыемхъ'
DEFINE BAR 18 OF popS7 PROMPT '\-'
DEFINE BAR 19 OF popS7 PROMPT 'тнплю хт-01'
DEFINE BAR 20 OF popS7 PROMPT 'тнплю нмй-01'
DEFINE BAR 21 OF popS7 PROMPT 'тнплю цн-01'
DEFINE BAR 22 OF popS7 PROMPT 'тнплю нмй-02'
DEFINE BAR 23 OF popS7 PROMPT 'тнплю нмй-03'
DEFINE BAR 24 OF popS7 PROMPT 'едхмюъ лерндхйю яюмйжхи'
DEFINE BAR 25 OF popS7 PROMPT 'лнмхрнпхмц щйн'
DEFINE BAR 26 OF popS7 PROMPT 'тнплю цн-02 (нмйнкнцхъ)'
DEFINE BAR 27 OF popS7 PROMPT 'нрвер он ощр'
ON SELECTION BAR 01 OF popS7 DO FormSh0
ON SELECTION BAR 02 OF popS7 DO FormSh1
ON SELECTION BAR 03 OF popS7 DO FormSh2
ON SELECTION BAR 04 OF popS7 DO FormSh3
ON SELECTION BAR 05 OF popS7 DO FormSh3Bis
ON SELECTION BAR 06 OF popS7 FormSh4(1) && лщщ
ON SELECTION BAR 07 OF popS7 FormSh4(2) && щйло
ON SELECTION BAR 08 OF popS7 DO FormSh5
ON SELECTION BAR 09 OF popS7 DO FormA30
ON SELECTION BAR 10 OF popS7 DO FormSh5Bis
ON SELECTION BAR 11 OF popS7 DO FormSh55Bis
ON SELECTION BAR 13 OF popS7 DO FormSh6
ON SELECTION BAR 14 OF popS7 do RepExp7
ON SELECTION BAR 15 OF popS7 do FormSh8
ON SELECTION BAR 16 OF popS7 do FormSh09
ON SELECTION BAR 17 OF popS7 do FormSh10
ON SELECTION BAR 19 OF popS7 do MakeIF01
ON SELECTION BAR 20 OF popS7 do FormOnk01
ON SELECTION BAR 21 OF popS7 do FormGOS701
ON SELECTION BAR 22 OF popS7 do FormGOS702
ON SELECTION BAR 23 OF popS7 do FormOnk03
ON SELECTION BAR 24 OF popS7 do yu_01
ON SELECTION BAR 25 OF popS7 do yu_02
ON SELECTION BAR 26 OF popS7 do yu_03
ON SELECTION BAR 27 OF popS7 do MakePet

DEFINE POPUP popS2 MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popS2 PROMPT 'ябндмши гю леяъж'
DEFINE BAR 02 OF popS2 PROMPT 'керюкэмше хяундш (онкмши)'
DEFINE BAR 03 OF popS2 PROMPT 'керюкэмше хяундш (йпюрйхи)'
ON SELECTION BAR 01 OF popS2 DO FormS20
ON SELECTION BAR 02 OF popS2 DO FormS21
ON SELECTION BAR 03 OF popS2 DO FormS22

DEFINE POPUP ExpCriteria MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1  OF ExpCriteria PROMPT 'керюкэмше хяундш'
DEFINE BAR 2  OF ExpCriteria PROMPT 'онбрнпмше цняохрюкхгюжхх'
DEFINE BAR 3  OF ExpCriteria PROMPT 'онбрнпмше цняохрюкхгюжхх (VER.2)'
DEFINE BAR 4  OF ExpCriteria PROMPT 'оепеяевемхъ бхднб ледонлных'
DEFINE BAR 5  OF ExpCriteria PROMPT 'нранп он опнтхкъл'
DEFINE BAR 6  OF ExpCriteria PROMPT '"всфхе" дхяоюмяепхгюжхх'
DEFINE BAR 7  OF ExpCriteria PROMPT '"всфхе" гюярпюунбюммше'
DEFINE BAR 8  OF ExpCriteria PROMPT '"всфхе" гюярпюунбюммше (VER.2)'
DEFINE BAR 9  OF ExpCriteria PROMPT 'цянохрюкхгюжхх <>50%'
DEFINE BAR 10 OF ExpCriteria PROMPT 'цянохрюкхгюжхх <>50% (VER.2)'
DEFINE BAR 11 OF ExpCriteria PROMPT '\-'
DEFINE BAR 12 OF ExpCriteria PROMPT 'цняохрюкхгюжхъ аег опедбюпхрекэмнцн наякеднбюмхъ'
ON SELECTION BAR 1 OF ExpCriteria do seldeads
ON SELECTION BAR 2 OF ExpCriteria do seldblgosps
ON SELECTION BAR 3 OF ExpCriteria do seldblgospsv2
ON SELECTION BAR 4 OF ExpCriteria do selcrosss
ON SELECTION BAR 5 OF ExpCriteria do selprofus
ON SELECTION BAR 6 OF ExpCriteria do seldsps
ON SELECTION BAR 7 OF ExpCriteria do SelGuests
ON SELECTION BAR 8 OF ExpCriteria do SelGuestsV2
ON SELECTION BAR 9 OF ExpCriteria do sel50P
ON SELECTION BAR 10 OF ExpCriteria do sel50PV2
ON SELECTION BAR 12 OF ExpCriteria do GospWOLech

DEFINE POPUP ExpCritNew MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF ExpCritNew PROMPT 'керюкэмше хяундш'
DEFINE BAR 2 OF ExpCritNew PROMPT 'онбрнпмше цняохрюкхгюжхх'
DEFINE BAR 3 OF ExpCritNew PROMPT 'оепеяевемхъ бхднб ледонлных' SKIP 
DEFINE BAR 4 OF ExpCritNew PROMPT 'нранп он опнтхкъл' SKIP 
DEFINE BAR 5 OF ExpCritNew PROMPT '"всфхе" дхяоюмяепхгюжхх' SKIP 
DEFINE BAR 6 OF ExpCritNew PROMPT '"всфхе" гюярпюунбюммше' SKIP 
DEFINE BAR 7 OF ExpCritNew PROMPT 'цянохрюкхгюжхх <>50%' SKIP 
DEFINE BAR 8 OF ExpCritNew PROMPT 'онбрнпмше напюыемхъ'
DEFINE BAR 9 OF ExpCritNew PROMPT '\-'
DEFINE BAR 10 OF ExpCritNew PROMPT 'гюцпсгхрэ пегскэрюрш щйяоепрхг'
ON SELECTION BAR 1 OF ExpCritNew do seldeadsnew
ON SELECTION BAR 2 OF ExpCritNew do seldblgospsnew
*ON SELECTION BAR 3 OF ExpCritNew do selcrosss
*ON SELECTION BAR 4 OF ExpCritNew do selprofus
*ON SELECTION BAR 5 OF ExpCritNew do seldsps
*ON SELECTION BAR 6 OF ExpCritNew do SelGuests
*ON SELECTION BAR 7 OF ExpCritNew do sel50P
ON SELECTION BAR 8 OF ExpCritNew do seldblobrsnew
ON SELECTION BAR 10 OF ExpCritNew do LoadResults

DEFINE POPUP popBuch MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1  OF popBuch PROMPT 'дхттепемжхпнбюммши ондсьебни мнплюрхб'
DEFINE BAR 2  OF popBuch PROMPT 'дхттепемжхпнбюммши ондсьебни мнплюрхб (ярнлюрнкнцхъ)'
DEFINE BAR 3  OF popBuch PROMPT '\-'
DEFINE BAR 4  OF popBuch PROMPT 'опнялнрп юоят' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 5  OF popBuch PROMPT 'опнялнрп опхкнфемхъ ╧4' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 6  OF popBuch PROMPT 'опнялнрп опхкнфемхъ ╧4 (ярнлюр)' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 7  OF popBuch PROMPT 'оевюрэ юоят (бюпхюмр 1)'  SKIP
DEFINE BAR 8  OF popBuch PROMPT 'оевюрэ юоят (бюпхюмр 2)'  SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 9  OF popBuch PROMPT 'юмюкхг бшонкмемхъ окюмнбшу назелнб'  SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 10  OF popBuch PROMPT '\-'
DEFINE BAR 11  OF popBuch PROMPT 'ятнплхпнбюрэ опхкнфемхе ╧4' SKIP
DEFINE BAR 12 OF popBuch PROMPT 'ятнплхпнбюрэ опхкнфемхе ╧4' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 13 OF popBuch PROMPT 'ятнплхпнбюрэ опхкнфемхе ╧4(ярнлюр)' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 14 OF popBuch PROMPT 'гюцпсгйю юбюмянб' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 15 OF popBuch PROMPT 'тнплхпнбюмхе опхкнфемхи ╧1'  SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 16 OF popBuch PROMPT '\-'
DEFINE BAR 17 OF popBuch PROMPT 'ябндмюъ беднлнярэ опхкнфемхъ 4 (VER.1, I3)'
*DEFINE BAR 18 OF popBuch PROMPT 'ябндмюъ беднлнярэ опхкнфемхъ 4 (VER.2)'
DEFINE BAR 18 OF popBuch PROMPT 'ябндмюъ беднлнярэ опхкнфемхъ 4 (VER.3, S7)'
DEFINE BAR 19 OF popBuch PROMPT 'ябндмюъ беднлнярэ опхкнфемхъ 4 (ярнлюр, VER.1, I3)'
*DEFINE BAR 21 OF popBuch PROMPT 'ябндмюъ беднлнярэ опхкнфемхъ 4 (ярнлюр, VER.2)'
DEFINE BAR 20 OF popBuch PROMPT 'ябндмюъ беднлнярэ опхкнфемхъ 4 (ярнлюр, VER.3, S7)'
DEFINE BAR 21 OF popBuch PROMPT '\-'
DEFINE BAR 22 OF popBuch PROMPT 'гюцпсгйю ME-тюикнб'
DEFINE BAR 23 OF popBuch PROMPT 'гюцпсгйю тюикнб хлонпрю'
DEFINE BAR 24 OF popBuch PROMPT 'опхкнфемхе й янцюгс ╧1'
DEFINE BAR 25 OF popBuch PROMPT '\-'
DEFINE BAR 26 OF popBuch PROMPT 'ябндмше пееярпш (янцюг-лед)'
DEFINE BAR 27 OF popBuch PROMPT 'тнплю люц-1 (хмцняярпюу-л)' SKIP 
DEFINE BAR 28 OF popBuch PROMPT 'тнплю люц-2 (хмцняярпюу-л)'
DEFINE BAR 29 OF popBuch PROMPT 'тнплю U-1 (янцюг)'
DEFINE BAR 30 OF popBuch PROMPT 'тнплю U-2 (янцюг)'
DEFINE BAR 31 OF popBuch PROMPT '\-'
DEFINE BAR 32 OF popBuch PROMPT '"хця-л" нрвер (щйяоепрхгю)'
DEFINE BAR 33 OF popBuch PROMPT '\-'
DEFINE BAR 34 OF popBuch PROMPT 'тнплю гог (рюакхжю 5)'
DEFINE BAR 35 OF popBuch PROMPT 'тнплю гог (рюакхжю 10)'
DEFINE BAR 36 OF popBuch PROMPT 'тнплю гйг ттнля'
DEFINE BAR 37 OF popBuch PROMPT '\-'
DEFINE BAR 38 OF popBuch PROMPT 'нрвер он онкс/бнгпюярс (демэцх)' 
DEFINE BAR 39 OF popBuch PROMPT 'нрвер он онкс/бнгпюярс (кчдх)' 
DEFINE BAR 40 OF popBuch PROMPT 'нрвер он онкс/бнгпюярс (демэцх) ярюж' 
DEFINE BAR 41 OF popBuch PROMPT 'нрвер он онкс/бнгпюярс (кчдх) ярюж' 
DEFINE BAR 42 OF popBuch PROMPT '\-'
DEFINE BAR 43 OF popBuch PROMPT 'нрвер он назелюл (repVolumes)'
DEFINE BAR 44 OF popBuch PROMPT 'нрвер йб (ттнля)'
DEFINE BAR 45 OF popBuch PROMPT 'нрвер ч-04'
DEFINE BAR 46 OF popBuch PROMPT 'нрвер ч-05'
DEFINE BAR 47 OF popBuch PROMPT 'нрвер ч-06'
DEFINE BAR 48 OF popBuch PROMPT 'нрвер ч-07'

ON SELECTION BAR 1 OF popBuch DO FORM ViewDifNorm
ON SELECTION BAR 2 OF popBuch DO FORM ViewDifNormS
ON SELECTION BAR 4 OF popBuch DO FORM ViewAPSF
*ON SELECTION BAR 5 OF popBuch DO FORM IIF(tdat1<{01.07.2014}, 'ViewPr4', 'Viewpr4n')
ON SELECTION BAR 5 OF popBuch DO FORM Viewpr4n
ON SELECTION BAR 6 OF popBuch DO FORM ViewPr4St
ON SELECTION BAR 7 OF popBuch DO MakeAPSF
ON SELECTION BAR 8 OF popBuch DO MakeAPSF2
ON SELECTION BAR 9 OF popBuch DO VolControls
ON SELECTION BAR 11 OF popBuch DO MakePr4
ON SELECTION BAR 12 OF popBuch DO MakePr4n
ON SELECTION BAR 13 OF popBuch DO MakePr4St
ON BAR 14 OF popBuch ACTIVATE POPUP popAvances
ON SELECTION BAR 15 OF popBuch DO MakePrilN1
ON SELECTION BAR 17 OF popBuch DO SvodPr4
*ON SELECTION BAR 18 OF popBuch DO SvodPr4V2
ON SELECTION BAR 18 OF popBuch DO SvodPr4V3
ON SELECTION BAR 19 OF popBuch DO SvodPr4St
*ON SELECTION BAR 21 OF popBuch DO SvodPr4StV2
ON SELECTION BAR 20 OF popBuch DO SvodPr4StV3
ON SELECTION BAR 22 OF popBuch DO LoadMeFiles
ON SELECTION BAR 23 OF popBuch DO LoadImpFiles
ON SELECTION BAR 24 OF popBuch DO Pril1S7
ON SELECTION BAR 26 OF popBuch DO SvRS7
ON SELECTION BAR 27 OF popBuch DO FormMag01
*ON SELECTION BAR 28 OF popBuch DO FormMag02
ON SELECTION BAR 28 OF popBuch DO FormMag02n
ON SELECTION BAR 29 OF popBuch DO FormU01
ON SELECTION BAR 30 OF popBuch DO FormU02
ON SELECTION BAR 32 OF popBuch DO MakeIGSM
ON SELECTION BAR 34 OF popBuch DO MakeZPZ
ON SELECTION BAR 35 OF popBuch DO MakeZPZT10
ON SELECTION BAR 36 OF popBuch DO MakeVKS
ON SELECTION BAR 38 OF popBuch DO SagOpl
ON SELECTION BAR 39 OF popBuch DO SagOpl2
ON SELECTION BAR 40 OF popBuch DO SagOpls
ON SELECTION BAR 41 OF popBuch DO SagOpl2s
ON SELECTION BAR 43 OF popBuch DO repVolumes
ON SELECTION BAR 44 OF popBuch DO rep_kv
ON SELECTION BAR 45 OF popBuch DO yu_04
ON SELECTION BAR 46 OF popBuch DO yu_05
ON SELECTION BAR 47 OF popBuch DO yu_06
ON SELECTION BAR 48 OF popBuch DO yu_07

DEFINE POPUP popAvances MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF popAvances PROMPT 'онксвеммши б нрвермнл оепхнде'
DEFINE BAR 2 OF popAvances PROMPT 'онксвеммши б рейсыел леяъже'
DEFINE BAR 3 OF popAvances PROMPT '\-'
DEFINE BAR 4 OF popAvances PROMPT 'гюцпсгйю ямърхи он 5.3.2'
DEFINE BAR 5 OF popAvances PROMPT 'гюцпсгйю назелнб'

ON SELECTION BAR 1 OF popAvances DO AvansPeriod
ON SELECTION BAR 2 OF popAvances DO AvansMonth 
ON SELECTION BAR 4 OF popAvances DO LoadS532
ON SELECTION BAR 5 OF popAvances DO LoadVols

DEFINE POPUP popTuneUp MARGIN RELATIVE SHADOW COLOR SCHEME 4
DEFINE BAR 1  OF popTuneUp PROMPT 'бшанп нрвермнцн оепхндю' 
DEFINE BAR 2  OF popTuneUp PROMPT '\-'
DEFINE BAR 3  OF popTuneUp PROMPT 'мюярпнийю пюанвху дхпейрнпхи'
DEFINE BAR 4  OF popTuneUp PROMPT '\-'
DEFINE BAR 5  OF popTuneUp PROMPT 'оепехмдейяюжхъ ад мях'
DEFINE BAR 6  OF popTuneUp PROMPT 'оепехмдейяюжхъ пюанвху аюг'
DEFINE BAR 7  OF popTuneUp PROMPT 'йнппейрхпнбйю ярпсйрспш пюанвху ад' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 8  OF popTuneUp PROMPT '\-'
DEFINE BAR 9  OF popTuneUp PROMPT 'союйнбюрэ тюикш ньханй' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 10 OF popTuneUp PROMPT 'намскхрэ тюикш ньханй' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 11 OF popTuneUp PROMPT 'сдюкхрэ CTRL-ЙХ' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 12 OF popTuneUp PROMPT 'сдюкхрэ тюикш нрбернб' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 13 OF popTuneUp PROMPT 'сдюкхрэ опнрнйнкш' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 14 OF popTuneUp PROMPT 'сдюкхрэ Mc-тюикш' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 15 OF popTuneUp PROMPT 'сдюкхрэ Mk-тюикш' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 16 OF popTuneUp PROMPT 'сдюкхрэ Mt-тюикш' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 17 OF popTuneUp PROMPT 'сдюкхрэ тюикш b_flk' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 18 OF popTuneUp PROMPT 'сдюкхрэ тюикш b_mek' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 19 OF popTuneUp PROMPT 'сдюкхрэ BAK-тюикш' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 20 OF popTuneUp PROMPT 'пюглнпнгхрэ лщй' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 21 OF popTuneUp PROMPT '\-'
DEFINE BAR 22 OF popTuneUp PROMPT '"янапюрэ" тюик рюпхтю' SKIP
DEFINE BAR 23 OF popTuneUp PROMPT 'союйнбюрэ ябндмше аюгш'  SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 24 OF popTuneUp PROMPT '\-'
DEFINE BAR 25 OF popTuneUp PROMPT 'юйрсюкхгхпнбюрэ мях'
DEFINE BAR 26 OF popTuneUp PROMPT 'люцхвеяйюъ ймнойю'
DEFINE BAR 27 OF popTuneUp PROMPT '\-'
DEFINE BAR 28 OF popTuneUp PROMPT 'йнмбепрхпнбюрэ тюик medicament.xml'
DEFINE BAR 29 OF popTuneUp PROMPT 'йнмбепрхпнбюрэ тюик medicament_man_pack.xml'
DEFINE BAR 30 OF popTuneUp PROMPT 'йнмбепрхпнбюрэ тюик medicament_mfc.xml'
DEFINE BAR 31 OF popTuneUp PROMPT '\-'
DEFINE BAR 32 OF popTuneUp PROMPT 'ябндмши нмйн-тюик гю оепхнд'
DEFINE BAR 33 OF popTuneUp PROMPT 'янапюрэ тюикш цняохрюкхгюжхи'

ON SELECTION BAR 1  OF popTuneUp DO FORM SetPeriod
ON SELECTION BAR 3  OF popTuneUp DO FORM TuneBase
ON SELECTION BAR 5  OF popTuneUp DO ComReind
ON SELECTION BAR 6  OF popTuneUp DO BasReind with m.gcPeriod
ON SELECTION BAR 7  OF popTuneUp DO CorStruct
ON SELECTION BAR 9  OF popTuneUp DO PackBD
ON SELECTION BAR 10 OF popTuneUp DO ZapEFiles
ON SELECTION BAR 11 OF popTuneUp DO DelAllCtrl
ON SELECTION BAR 12 OF popTuneUp DO DelAllZapros
ON SELECTION BAR 13 OF popTuneUp DO DelAllProtocols
ON SELECTION BAR 14 OF popTuneUp DO DelMcFiles
ON SELECTION BAR 15 OF popTuneUp DO DelMkFiles
ON SELECTION BAR 16 OF popTuneUp DO DelMtFiles
ON SELECTION BAR 17 OF popTuneUp DO DelAllBFlk
ON SELECTION BAR 18 OF popTuneUp DO DelAllBMek
ON SELECTION BAR 19 OF popTuneUp DO DelBakFiles
ON SELECTION BAR 20 OF popTuneUp DO DeFrMek
*ON SELECTION BAR 22 OF popTuneUp DO MakeTarif
ON SELECTION BAR 23 OF popTuneUp DO PackBDSv
ON SELECTION BAR 25 OF popTuneUp DO ActNSI
ON SELECTION BAR 26 OF popTuneUp DO PushMagButton
ON SELECTION BAR 28 OF popTuneUp DO Medicament_sax
ON SELECTION BAR 29 OF popTuneUp DO MedPack_sax
ON SELECTION BAR 30 OF popTuneUp DO MedMFC
ON SELECTION BAR 32 OF popTuneUp DO MakeOnkoFile
ON SELECTION BAR 33 OF popTuneUp DO MakeAllGosps

DEFINE POPUP popMedSpr MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF popMedSpr PROMPT 'рюпхт'
DEFINE BAR 2 OF popMedSpr PROMPT 'лйа-10'
DEFINE BAR 3 OF popMedSpr PROMPT 'мнплюрхбмше назелш сяксц '
DEFINE BAR 4 OF popMedSpr PROMPT 'опнтхкх ледхжхмяйни онлных'

IF m.IsNotePad = .F.
 ON SELECTION BAR 1 OF popMedSpr DO FORM TarifN
 ON SELECTION BAR 2 OF popMedSpr DO FORM Mkb10
 ON SELECTION BAR 3 OF popMedSpr DO FORM CodKU
ELSE 
 ON SELECTION BAR 1 OF popMedSpr DO FORM TarifN600
 ON SELECTION BAR 2 OF popMedSpr DO FORM Mkb10600
 ON SELECTION BAR 3 OF popMedSpr DO FORM CodKU600
ENDIF 
ON SELECTION BAR 4 OF popMedSpr DO FORM ViewLicAll

DEFINE POPUP popTuneUp2 MARGIN RELATIVE SHADOW COLOR SCHEME 4
DEFINE BAR 1  OF popTuneUp2 PROMPT 'йнппейрхпнбйю ябндмни аюгш' 
DEFINE BAR 2  OF popTuneUp2 PROMPT 'намскхрэ тюикш щйяоепрхгш' 
DEFINE BAR 3  OF popTuneUp2 PROMPT '\-' 
DEFINE BAR 4  OF popTuneUp2 PROMPT 'сдюкемхе меопнтхкэмши детейрнб щйло' 
DEFINE BAR 5  OF popTuneUp2 PROMPT 'опнбепйю RID' SKIP 
DEFINE BAR 6  OF popTuneUp2 PROMPT '\-' 
DEFINE BAR 7  OF popTuneUp2 PROMPT '"опнюодеирхрэ" мюгбюмхъ б рюпхте' 
DEFINE BAR 8  OF popTuneUp2 PROMPT 'опнбепйю смхйюкэмнярх M-тюикнб' 
DEFINE BAR 9  OF popTuneUp2 PROMPT '\-' 
DEFINE BAR 10 OF popTuneUp2 PROMPT 'хлонпр ябндмни ад б MSSQL' 
DEFINE BAR 11 OF popTuneUp2 PROMPT 'сдюкхрэ тюикш SOAP' 
DEFINE BAR 12 OF popTuneUp2 PROMPT 'янгдюрэ яопюбнвмхйх кхжемгхи' 
DEFINE BAR 13  OF popTuneUp2 PROMPT '\-' 
DEFINE BAR 14  OF popTuneUp2 PROMPT 'сдюкхрэ оепянрвер' 
DEFINE BAR 15  OF popTuneUp2 PROMPT 'бняярюмнбхрэ тюикш ньханй он оепянрверш' 
DEFINE BAR 16  OF popTuneUp2 PROMPT 'ярюрхярхйю гюонкмемхъ тюикнб' 
DEFINE BAR 17  OF popTuneUp2 PROMPT 'пюяявхрюрэ пюглеп дхпейрнпхи' 
DEFINE BAR 18  OF popTuneUp2 PROMPT 'сдюкхрэ XLS-тюикш' 
DEFINE BAR 19  OF popTuneUp2 PROMPT '\-' 
DEFINE BAR 20  OF popTuneUp2 PROMPT 'ярюрхярхйю ямърхи' 
DEFINE BAR 21  OF popTuneUp2 PROMPT 'FOXCHART' 
DEFINE BAR 22  OF popTuneUp2 PROMPT 'PDFCREATOR' 
DEFINE BAR 23  OF popTuneUp2 PROMPT '\-' 
DEFINE BAR 24  OF popTuneUp2 PROMPT 'хлонпр бра' 
DEFINE BAR 25  OF popTuneUp2 PROMPT 'йнмянкхдюжхъ аюг MEE (хмцняярпюу-л)' 
DEFINE BAR 26  OF popTuneUp2 PROMPT 'йнмянкхдюжхъ аюг BASE (хмцняярпюу-л)' 
DEFINE BAR 27  OF popTuneUp2 PROMPT 'йнмянкхдюжхъ аюг BASE (2-ни щрюо, хмцняярпюу-л)' 
DEFINE BAR 28  OF popTuneUp2 PROMPT 'сдюкемхе дсакеи хг M-тюикнб (хмцняярпюу-л)'  SKIP 
DEFINE BAR 29  OF popTuneUp2 PROMPT '\-' 
DEFINE BAR 30  OF popTuneUp2 PROMPT 'янапюрэ ябндмши CTRL' 
DEFINE BAR 31  OF popTuneUp2 PROMPT 'япюбмхрэ дбю CTRL' 
DEFINE BAR 32  OF popTuneUp2 PROMPT 'оепедюрэ ньхайх' 
DEFINE BAR 33  OF popTuneUp2 PROMPT '\-' 
DEFINE BAR 34  OF popTuneUp2 PROMPT 'янапюрэ ME-тюикш' 
DEFINE BAR 35  OF popTuneUp2 PROMPT 'сдюкхрэ hosp-тюикш' 
DEFINE BAR 36  OF popTuneUp2 PROMPT 'бйкчвхрэ рюилеп оепехмдейяюжхх' 

ON SELECTION BAR 1  OF popTuneUp2 Do CorSvBases
ON SELECTION BAR 2  OF popTuneUp2 Do KillMeFiles
ON SELECTION BAR 4  OF popTuneUp2 Do KillBadEkmp
*ON SELECTION BAR 5  OF popTuneUp2 Do CheckRID
ON SELECTION BAR 7  OF popTuneUp2 Do UpdNamesInTarif
ON SELECTION BAR 8  OF popTuneUp2 Do CleanMFiles
ON SELECTION BAR 11 OF popTuneUp2 Do KillSoapFiles
ON SELECTION BAR 12 OF popTuneUp2 CreateLicences()
ON SELECTION BAR 14 OF popTuneUp2 do DelYFiles
ON SELECTION BAR 15 OF popTuneUp2 do RestEFls2
ON SELECTION BAR 16 OF popTuneUp2 do StatFillFiles
ON SELECTION BAR 17 OF popTuneUp2 do DirSize
ON SELECTION BAR 18 OF popTuneUp2 do DelXlsFiles
ON SELECTION BAR 20 OF popTuneUp2 do MekDefsStat
ON SELECTION BAR 21 OF popTuneUp2 do myChart
ON SELECTION BAR 22 OF popTuneUp2 do pdf_test_002
ON SELECTION BAR 24 OF popTuneUp2 do Imp2R2
ON SELECTION BAR 25 OF popTuneUp2 do ConsMEE
ON SELECTION BAR 26 OF popTuneUp2 do ConsBase
ON SELECTION BAR 27 OF popTuneUp2 do ConsBase2
ON SELECTION BAR 28 OF popTuneUp2 do DelDblsMFiles
ON SELECTION BAR 30 OF popTuneUp2 do MakeSVCtrl
ON SELECTION BAR 31 OF popTuneUp2 do Cmp2Ctrls2
ON SELECTION BAR 32 OF popTuneUp2 do PassEX
ON SELECTION BAR 34 OF popTuneUp2 do SumMeFiles
ON SELECTION BAR 35 OF popTuneUp2 do DelHospFiles
ON SELECTION BAR 36 OF popTuneUp2 do ReindexTimer

DEFINE POPUP popPostgreSQL MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popPostgreSQL PROMPT 'хлонпр б PostgreSQL' PICTURE 'POSTGRESQL.GIF'
DEFINE BAR 02 OF popPostgreSQL PROMPT '\-'
DEFINE BAR 03 OF popPostgreSQL PROMPT 'хлонпр пюглепмняреи б MS SQL'
DEFINE BAR 04 OF popPostgreSQL PROMPT 'хлонпр пюанвеи ад б MS SQL'

DEFINE BAR 05 OF popPostgreSQL PROMPT 'хлонпр AISOMS б MS SQL'
DEFINE BAR 06 OF popPostgreSQL PROMPT 'хлонпр PR4 б MS SQL'
DEFINE BAR 07 OF popPostgreSQL PROMPT 'хлонпр PR4ST б MS SQL'
DEFINE BAR 08 OF popPostgreSQL PROMPT 'хлонпр MAG02 б MS SQL'

DEFINE BAR 09 OF popPostgreSQL PROMPT 'хлонпр HO б MS SQL'
DEFINE BAR 10 OF popPostgreSQL PROMPT 'хлонпр PEOPLE б MS SQL'
DEFINE BAR 11 OF popPostgreSQL PROMPT 'хлонпр OTDEL б MS SQL'
DEFINE BAR 12 OF popPostgreSQL PROMPT 'хлонпр лщй б MS SQL'
DEFINE BAR 13 OF popPostgreSQL PROMPT 'онхяй осяршу SQLID' SKIP 
DEFINE BAR 14 OF popPostgreSQL PROMPT '\-'
DEFINE BAR 15 OF popPostgreSQL PROMPT 'оепеявхрюрэ Typ, Mp, Vz' SKIP 
DEFINE BAR 16 OF popPostgreSQL PROMPT 'тнплхпнбюмхе UDST-тюикю'
DEFINE BAR 17 OF popPostgreSQL PROMPT '\-'
DEFINE BAR 18 OF popPostgreSQL PROMPT 'хлонпр пюанвеи ад б CSV'
DEFINE BAR 19 OF popPostgreSQL PROMPT 'оепеявхрюрэ назелш'
DEFINE BAR 20 OF popPostgreSQL PROMPT 'яанп ANSWERS'
DEFINE BAR 21 OF popPostgreSQL PROMPT 'хцпю я гюопняюлх'

ON SELECTION BAR 01 OF popPostgreSQL Do Lpu2Postgre
ON SELECTION BAR 03 OF popPostgreSQL Do Dims2SQL
ON SELECTION BAR 04 OF popPostgreSQL Do Fact2SQL

ON SELECTION BAR 05 OF popPostgreSQL Do AisOms2SQL
ON SELECTION BAR 06 OF popPostgreSQL Do Pr42SQL
ON SELECTION BAR 07 OF popPostgreSQL Do Pr4St2SQL
ON SELECTION BAR 08 OF popPostgreSQL Do Mag2SQL

ON SELECTION BAR 09 OF popPostgreSQL Do HO2SQL
ON SELECTION BAR 10 OF popPostgreSQL Do People2SQL
ON SELECTION BAR 11 OF popPostgreSQL Do Otdel2SQL
ON SELECTION BAR 12 OF popPostgreSQL Do Mek2SQL
ON SELECTION BAR 13 OF popPostgreSQL Do FEmptySqlId
ON SELECTION BAR 15 OF popPostgreSQL Do AllMpTyp
ON SELECTION BAR 16 OF popPostgreSQL Do MakeUDSt
ON SELECTION BAR 18 OF popPostgreSQL Do Fact2CSV
ON SELECTION BAR 19 OF popPostgreSQL Do IIF(INLIST(m.qcod,'S7','R2'), 're_vols_s7', 're_volumes')
ON SELECTION BAR 20 OF popPostgreSQL Do SumAnswers
ON SELECTION BAR 21 OF popPostgreSQL Do PlayRequests

DEFINE POPUP popSOAP MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popSOAP PROMPT 'SOAP (MSOSOAP.SOAPClient)' SKIP 
DEFINE BAR 02 OF popSOAP PROMPT 'findPersonByPolicy'
DEFINE BAR 03 OF popSOAP PROMPT 'findPersons'
DEFINE BAR 04 OF popSOAP PROMPT 'getPersonPolicy'
DEFINE BAR 05 OF popSOAP PROMPT 'getPersonsPolicy'
DEFINE BAR 06 OF popSOAP PROMPT '\-'
DEFINE BAR 07 OF popSOAP PROMPT 'GetPersonInsuranceDataAsync'
DEFINE BAR 08 OF popSOAP PROMPT 'GetPersonInsuranceDataAsync (ver. 01)'
DEFINE BAR 09 OF popSOAP PROMPT 'GetPersonInsuranceDataMassAsync'
DEFINE BAR 10 OF popSOAP PROMPT '\-'
DEFINE BAR 11 OF popSOAP PROMPT 'getBillStatuses (БЯЕ лн)'
DEFINE BAR 12 OF popSOAP PROMPT 'getBillStatuses (лн 1798, 0207047)'
DEFINE BAR 13 OF popSOAP PROMPT 'getMailGw'
DEFINE BAR 14 OF popSOAP PROMPT 'getPdf'
DEFINE BAR 15 OF popSOAP PROMPT 'getAttachment'
DEFINE BAR 16 OF popSOAP PROMPT 'changeBillStatus'
DEFINE BAR 17 OF popSOAP PROMPT 'uploadMail'
DEFINE BAR 18 OF popSOAP PROMPT '\-'
DEFINE BAR 19 OF popSOAP PROMPT 'getDictionaries'
DEFINE BAR 20 OF popSOAP PROMPT 'getDictionary'
DEFINE BAR 21 OF popSOAP PROMPT '\-'
DEFINE BAR 22 OF popSOAP PROMPT 'getXmlAttachment'
DEFINE BAR 23 OF popSOAP PROMPT 'пюяоюпяхрэ тедепюкэмши апед'

ON SELECTION BAR 01 OF popSOAP do soap01
ON SELECTION BAR 02 OF popSOAP do findPersonByPolicy
ON SELECTION BAR 03 OF popSOAP do findPersons
ON SELECTION BAR 04 OF popSOAP do getPersonPolicy
ON SELECTION BAR 05 OF popSOAP do getPersonsPolicy
ON SELECTION BAR 07 OF popSOAP do GetPersonInsDataAsync
ON SELECTION BAR 08 OF popSOAP do GetPersonInsDataAsyncV01
ON SELECTION BAR 09 OF popSOAP do GetPersonInsDataMassAsync
ON SELECTION BAR 11 OF popSOAP getBillStatuses(0, null, .f., 'SMO')
ON SELECTION BAR 12 OF popSOAP getBillStatuses(1798, null, .f.)
ON SELECTION BAR 13 OF popSOAP getMailGw(0, null, "")
ON SELECTION BAR 14 OF popSOAP do getPdf
ON SELECTION BAR 15 OF popSOAP getAttachment("")
ON SELECTION BAR 16 OF popSOAP do changeBillStatus
ON SELECTION BAR 17 OF popSOAP uploadMail('1621406') && m.parentMailGWlogId, lpu_id=2295, mcod=4344931
ON SELECTION BAR 19 OF popSOAP getDictionaries()
ON SELECTION BAR 20 OF popSOAP getDictionary('sprul.00')
ON SELECTION BAR 22 OF popSOAP XmlAttTest()
ON SELECTION BAR 23 OF popSOAP Parse79()

DEFINE POPUP popParallel MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popParallel PROMPT 'INSTALL MULTITHREADING'
DEFINE BAR 02 OF popParallel PROMPT 'UNINSTALL MULTITHREADING'
DEFINE BAR 03 OF popParallel PROMPT '\-'
DEFINE BAR 04 OF popParallel PROMPT 'рЕЯР (BaseReindexParallel)'
*ON SELECTION BAR 01 OF popParallel do Install
*ON SELECTION BAR 02 OF popParallel do Uninstall
ON SELECTION BAR 04 OF popParallel DO BaseReindexParallel

SET SYSMENU AUTOMATIC
SET SYSMENU ON
