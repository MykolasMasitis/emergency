FUNCTION OneReind(lcPath)

tc_mcod = SUBSTR(lcPath, RAT('\', lcPath)+1)

IF OpenFile("&lcPath\Calls", "Calls", "excl") == 0
 SELECT Calls
 DELETE TAG ALL 
 SET FULLPATH OFF 
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON n_pol TAG n_pol
 INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
 INDEX ON dr TAG dr
 SET FULLPATH OFF 
 USE IN Calls 
ENDIF

IF OpenFile("&lcPath\Teams", "Teams", "excl") == 0
 SELECT Teams
 DELETE TAG ALL 
 SET FULLPATH OFF 
 INDEX ON c_br TAG c_br
 INDEX ON pst TAG pst
 SET FULLPATH OFF 
 USE IN Teams 
ENDIF

RETURN 
