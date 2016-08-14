FUNCTION SendToLpuPr(lcPath, cmcod, clpuid)

 m.mmy = PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)
 m.ctrl = 'ctrl'+m.mmy+'.dbf'
 m.prsp = 'prsp'+m.qcod+m.mmy+'.pdf'
 IF !fso.FileExists(lcpath+'\'+m.ctrl)
  MESSAGEBOX(CHR(13)+CHR(10)+'���� ������ �� �����������!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(lcpath+'\'+m.prsp)
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� �� �����������!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 ZipFile = 'pr'+clpuid+m.qcod+'.zip'
 IF fso.FileExists(lcPath+'\'+ZipFile)
  fso.DeleteFile(lcPath+'\'+ZipFile)
 ENDIF 
 
 ZipOpen(lcPath+'\'+ZipFile)
 ZipFile(lcPath+'\'+m.ctrl)
 ZipFile(lcPath+'\'+m.prsp)
 ZipClose()
 
 IF !fso.FileExists(lcPath+'\'+ZipFile)
  MESSAGEBOX(CHR(13)+CHR(10)+'���������� ������� �����!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 m.cto      = IIF(SEEK(INT(VAL(clpuid)), 'sprabo', 'lpu_id'), 'oms@'+sprabo.abn_name, '')

 m.un_id    = SYS(3)
 m.bansfile = 'b' + m.un_id
 m.tansfile = 't' + m.un_id
 m.dfile    = 'd' + m.un_id
 m.mmid     = m.un_id+'.'+m.gcUser+'@'+m.qmail
 m.csubj    = 'OMS#'+gcPeriod+'###SMPM'

 poi = fso.CreateTextFile(lcPath + '\' + m.tansfile)

 poi.WriteLine('To: '+m.cTO)
 poi.WriteLine('Message-Id: ' + m.mmid)
 poi.WriteLine('Subject: ' + m.csubj)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Resent-Message-Id: '+ALLTRIM(cmessage))
 poi.WriteLine('Attachment: '+m.dfile+' '+ZipFile)
 
 poi.Close
 
 fso.CopyFile(lcPath+'\'+ZipFile, pAisOms+'\oms\output\'+m.dfile)
 fso.CopyFile(lcPath+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)

 REPLACE issent WITH .t. 
 MESSAGEBOX('����� ����������!',0+64,'')

RETURN  