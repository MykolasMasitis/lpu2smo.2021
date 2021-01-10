FUNCTION dt2date(para1) && Конверитруем из [Thu,  2 Feb 2012 08:42:29 +0300] (RFC822) в datetime-формат
 PRIVATE lcData
 m.lcData = para1 && Thu,  2 Feb 2012 08:42:29 +0300
 
 IF SET("Hours")!=24
  SET HOURS TO 24
 ENDIF 
 m.lcdata = ALLTRIM(SUBSTR(m.lcdata, AT(',', m.lcdata)+1)) && 2 Feb 2012 08:42:29 +0300

 * startpos = RAT(' ',lcData,5)
 *lcDay    = PADL(ALLTRIM(SUBSTR(lcData,startpos,2)),2,'0')
 lcDay    = PADL(ALLTRIM(SUBSTR(m.lcdata, 1, AT(' ',m.lcdata)-1)),2,'0')
 *lcMonthT = SUBSTR(lcData,startpos+3,3)
 lcMonthT = SUBSTR(m.lcdata, AT(' ',m.lcdata,1)+1, AT(' ',m.lcdata,2)-(AT(' ',m.lcdata,1)+1))
 DO CASE
  CASE lcMonthT = 'Jan'
   lcMonth = '01'
  CASE lcMonthT = 'Feb'
   lcMonth = '02'
  CASE lcMonthT = 'Mar'
   lcMonth = '03'
  CASE lcMonthT = 'Apr'
   lcMonth = '04'
  CASE lcMonthT = 'May'
   lcMonth = '05'
  CASE lcMonthT = 'Jun'
   lcMonth = '06'
  CASE lcMonthT = 'Jul'
   lcMonth = '07'
  CASE lcMonthT = 'Aug'
   lcMonth = '08'
  CASE lcMonthT = 'Sep'
   lcMonth = '09'
  CASE lcMonthT = 'Oct'
   lcMonth = '10'
  CASE lcMonthT = 'Nov'
   lcMonth = '11'
  CASE lcMonthT = 'Dec'
   lcMonth = '12'
  OTHERWISE 
   lcMonth = '00'
 ENDCASE 

 *lcYear  =  SUBSTR(lcData, startpos+7, 4)
 lcYear  =  SUBSTR(m.lcdata, AT(' ',m.lcdata,2)+1, AT(' ',m.lcdata,3)-(AT(' ',m.lcdata,2)+1))
 lcDate  = lcDay +'.' + lcMonth + '.' + lcYear
 *lcTime = SUBSTR(lcData, startpos+12, 8)
 lcTime = SUBSTR(m.lcdata, AT(' ',m.lcdata,3)+1, AT(' ',m.lcdata,4)-(AT(' ',m.lcdata,3)+1))
 lcRealData = CTOT(lcDate + ' ' + lcTime)

RETURN lcRealData
