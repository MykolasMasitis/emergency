PROCEDURE PackBd
 IF MESSAGEBOX('��� ��������� ��������� �������'+CHR(13)+CHR(10)+;
  '���������� � �������� ������ ������ ������.'+CHR(13)+CHR(10)+'����������?',4+32, '')==7
  RETURN 
 ENDIF 
 
 IF OpenFile("&pBase\&gcPeriod\aisoms", "aisoms", "shar", "mcod") > 0
  RETURN
 ENDIF 
 IF OpenFile(pcommon+'\UsrLpu', "UsrLpu", "shar", "mcod") > 0
  USE IN aisoms
  RETURN
 ENDIF 

 SELECT AisOms
 
 SCAN
  m.mcod = mcod
  m.usr  = IIF(SEEK(m.mcod, "usrlpu"), 'USR'+PADL(usrlpu.usr,3,'0'), "")
  IF m.usr != m.gcUser AND m.gcUser!='OMS'
   LOOP 
  ENDIF 

  WAIT m.mcod WINDOW NOWAIT 
  
  IF !fso.FolderExists(pBase+'\'+gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  
  IF !fso.FileExists(pBase+'\'+gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'error', 'excl')>0
   LOOP 
  ENDIF 
  
  SELECT error
  PACK 
  USE 

 ENDSCAN 

 WAIT CLEAR 
 USE 
 USE IN UsrLpu
 
RETURN 