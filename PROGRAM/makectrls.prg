PROCEDURE MakeCTRLs
 IF !fso.FileExists(ptempl+'\ctrlmmy.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер ьюакнм тюикю ньханй CTRLMMY.DBF'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 IF OpenFile("&pBase\&gcPeriod\aisoms", "aisoms", "shar") > 0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile("&pBase\&gcPeriod\NSI\errors", "errors", "shar", "code") > 0
  IF USED('errors')
   USE IN errors
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 
 
 SCAN
  WAIT mcod WINDOW NOWAIT 

  =MakeCTRL(pbase+'\'+gcPeriod+'\'+mcod, mcod)
  
  SELECT AisOms

  IF CHRSAW(0) == .T.
   IF INKEY() == 27
    IF MESSAGEBOX('бш унрхре опепбюрэ напюанрйс?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDSCAN 

 USE IN aisoms
 USE IN errors

 WAIT CLEAR 

 SET ESCAPE &OldEscStatus

 MESSAGEBOX('напюанрйю гюйнмвемю!', 0+64, '')

RETURN 

