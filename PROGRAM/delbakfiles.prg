PROCEDURE DelBakFiles
 IF MESSAGEBOX('����� ������� ��� BAK-�����?'+CHR(13)+CHR(10)+;
               '��� ��, ��� �� ������������� ������ �������?',4+48,'') != 6
  RETURN 
 ENDIF 

 IF MESSAGEBOX('�� ��������� ������� � ����� ���������?',4+48,'') != 6
  RETURN 
 ENDIF 
 
 IF OpenFile("&pBase\&gcPeriod\aisoms", "aisoms", "shar", "mcod") > 0
  RETURN
 ENDIF 
 
 SELECT AisOms
 
 SCAN
  m.mcod = mcod

  WAIT m.mcod WINDOW NOWAIT 

  IF fso.FileExists(pBase+'\'+m.gcperiod+'\'+mcod+'\people.bak')
   fso.DeleteFile(pBase+'\'+m.gcperiod+'\'+mcod+'\people.bak')
  ENDIF 
  IF fso.FileExists(pBase+'\'+m.gcperiod+'\'+mcod+'\talon.bak')
   fso.DeleteFile(pBase+'\'+m.gcperiod+'\'+mcod+'\talon.bak')
  ENDIF 
  IF fso.FileExists(pBase+'\'+m.gcperiod+'\'+mcod+'\otdel.bak')
   fso.DeleteFile(pBase+'\'+m.gcperiod+'\'+mcod+'\otdel.bak')
  ENDIF 
  IF fso.FileExists(pBase+'\'+m.gcperiod+'\'+mcod+'\doctor.bak')
   fso.DeleteFile(pBase+'\'+m.gcperiod+'\'+mcod+'\doctor.bak')
  ENDIF 

 ENDSCAN 
 WAIT CLEAR 
 USE 
 USE IN UsrLpu

RETURN 