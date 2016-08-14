FUNCTION dt2date(lcData) && Конверитруем из [Thu,  2 Feb 2012 08:42:29 +0300] (RFC822) в datetime-формат
 IF SET("Hours")!=24
  SET HOURS TO 24
 ENDIF 

 lcDay    = PADL(ALLTRIM(SUBSTR(lcData,6,2)),2,'0')
 lcMonthT = SUBSTR(lcData,9,3)
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
 lcYear  =  SUBSTR(lcData, 13, 4)

 lcDate  = lcDay +'.' + lcMonth + '.' + lcYear
  
 lcTime = SUBSTR(lcData, 18, 8)
 
 lcGrinv = VAL(SUBSTR(lcData, 27, 3))
 
 lcRealData = IIF(lcGrinv != -8, CTOT(lcDate + ' ' + lcTime) - (lcGrinv*60*60) + 4*60*60, CTOT(lcDate + ' ' + lcTime)) && Прибавляем 4 часа к Гринвичу - Москва

RETURN lcRealData