PROCEDURE AllLocExp

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 
 
 SCAN
  WAIT mcod WINDOW NOWAIT 
  MailView.refresh
  
  m.mcod = mcod
  =LocExp(m.pbase+'\'+m.gcPeriod+'\'+m.mcod, m.mcod)
  
  SELECT AisOms

  IF CHRSAW(0) == .T.
   IF INKEY() == 27
    IF MESSAGEBOX('�� ������ �������� ���������?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDSCAN 

 WAIT CLEAR 

 SET ESCAPE &OldEscStatus

 MESSAGEBOX('��������� ���������!', 0+64, '')

RETURN 

