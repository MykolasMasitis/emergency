PROCEDURE ComReind

WAIT "оЕПЕХМДЕЙЮЖХЪ COMMON..." WINDOW NOWAIT 

IF OpenFile(pCommon+'\UsrLpu', "UsrLpu", "excl") == 0
 SELECT UsrLpu
 SET FULLPATH OFF 
 WAIT "хмдейяхпнбюмхе тюикю "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON mcod TAG mcod
 INDEX ON lpu_id TAG lpu_id
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF USED('Users')
 USE IN users
ENDIF 
IF OpenFile(pcommon+'\Users', 'users', 'excl') <= 0
 SELECT Users 
 INDEX ON name TAG name 
 USE 
ENDIF 
IF !USED('Users')
 =OpenFile(pCommon+'\Users', 'Users', 'shar', 'name')
ENDIF 

IF !fso.FileExists(pCommon+'\Users.cdx')
 IF OpenFile(pcommon+'\Users', 'users', 'excl') <= 0
  SELECT Users 
  INDEX ON name TAG name 
  USE 
 ENDIF 
ENDIF 

IF OpenFile(pCommon+'\emails', "emails", "excl") == 0
 SELECT emails
 SET FULLPATH OFF 
 WAIT "хмдейяхпнбюмхе тюикю "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpu_id TAG lpu_id
 INDEX ON mcod TAG mcod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pCommon+'\Territ', "territ", "excl") == 0
 SELECT territ
 SET FULLPATH OFF 
 WAIT "хмдейяхпнбюмхе тюикю "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cod TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pCommon+'\Mkb10', "mkb", "excl") == 0
 SELECT Mkb
 SET FULLPATH OFF 
 WAIT "хмдейяхпнбюмхе тюикю "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON ds TAG ds
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

WAIT CLEAR 

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\osoerzxx', "OsoERZ", "excl") == 0
 SELECT OsoERZ
 SET FULLPATH OFF 
 WAIT "хмдейяхпнбюмхе тюикю "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON Ans_r TAG Ans_r
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\smo', "smo", "excl") == 0
 SELECT smo
 SET FULLPATH OFF 
 WAIT "хмдейяхпнбюмхе тюикю "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON code TAG code
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spraboxx', "sprabo", "excl") == 0
 SELECT sprabo
 SET FULLPATH OFF 
 WAIT "хмдейяхпнбюмхе тюикю "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON abn_name TAG abn_name
 INDEX ON object_id TAG lpu_id
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "excl") == 0
 SELECT sprlpu
 SET FULLPATH OFF 
 WAIT "хмдейяхпнбюмхе тюикю "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpu_id TAG lpu_id
 INDEX ON fil_id TAG fil_id
 INDEX ON mcod TAG mcod
 INDEX ON cokr TAG cokr
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\tarifn', "tarif", "excl") == 0
 SELECT tarif
 SET FULLPATH OFF 
 WAIT "хмдейяхпнбюмхе тюикю "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cod TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\rsv009xx', "rsv", "excl") == 0
 SELECT rsv
 SET FULLPATH OFF 
 WAIT "хмдейяхпнбюмхе тюикю "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON rslt TAG rslt
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF !fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\errors.dbf')
 IF fso.FileExists(pcommon+'\errors.dbf')
  fso.CopyFile(pcommon+'\errors.dbf', pbase+'\'+gcperiod+'\'+'nsi'+'\errors.dbf')
 ENDIF 
ENDIF 

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\errors', "errors", "excl") == 0
 SELECT errors
 SET FULLPATH OFF 
 WAIT "хмдейяхпнбюмхе тюикю "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON code TAG code
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF
