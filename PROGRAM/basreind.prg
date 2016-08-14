FUNCTION BasReind

WAIT "Переиндексация aisoms.dbf..." WINDOW NOWAIT 
IF OpenFile(pBase+'\'+gcPeriod+'\AisOms', 'AisOms', 'excl') == 0
 SELECT AisOms
 DELETE TAG ALL 
 INDEX ON cmessage TAG cmessage
 INDEX ON mcod TAG mcod
 INDEX ON lpuid TAG lpuid
 INDEX ON TTOC(sent,1)      TAG sent 
 INDEX ON TTOC(recieved,1)  TAG recieved
 INDEX ON TTOC(processed,1) TAG processed
 INDEX ON usr TAG usr
 INDEX ON paztot TAG paztot
 USE 
ENDIF 
WAIT CLEAR 

WAIT "Переиндексация aisoms.dbf..." WINDOW NOWAIT 
IF OpenFile(pDouble+'\AisOms', 'AisOms', 'excl') == 0
 SELECT AisOms
 DELETE TAG ALL 
 INDEX ON cmessage TAG cmessage
 INDEX ON mcod TAG mcod
 INDEX ON lpuid TAG lpuid
 INDEX ON TTOC(sent,1)      TAG sent 
 INDEX ON TTOC(recieved,1)  TAG recieved
 INDEX ON TTOC(processed,1) TAG processed
 USE
ENDIF 
WAIT CLEAR 

WAIT "Переиндексация aisoms.dbf..." WINDOW NOWAIT 
IF OpenFile(pTrash+'\AisOms', 'AisOms', 'excl') == 0
 SELECT AisOms
 DELETE TAG ALL 
 INDEX ON cmessage TAG cmessage
 INDEX ON mcod TAG mcod
 INDEX ON lpuid TAG lpuid
 INDEX ON TTOC(sent,1)      TAG sent 
 INDEX ON TTOC(recieved,1)  TAG recieved
 INDEX ON TTOC(processed,1) TAG processed
 USE
ENDIF 
WAIT CLEAR 

tn_result = 0
tn_result = tn_result + OpenFile(pBase+'\'+gcPeriod+'\AisOms', 'AisOms', 'shar', 'mcod')
IF tn_result > 0
 RETURN 
ENDIF 

SELECT AisOms
SCAN  
 WAIT mcod WINDOW NOWAIT 
 lcPath = pbase+'\'+m.gcperiod+'\'+mcod
 IF fso.FileExists(lcPath+'\People.dbf')
  =OneReind(lcPath)
 ENDIF 
 WAIT CLEAR 
ENDSCAN 
USE 

RETURN 

