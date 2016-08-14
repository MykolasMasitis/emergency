FUNCTION Snd2MgfOMS

 m.mcod = '0371001'
 m.lcpath = m.pbase+'\'+m.gcperiod+'\'+m.mcod
 m.mmy   = PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)
 m.ctrl  = 'ctrl'+m.mmy+'.dbf'
 m.ctrln = 'ctrl'+m.qcod+m.mmy+'.dbf'
 m.prsp  = 'pmsp'+m.qcod+m.mmy+'.pdf'
 m.ssp   = 'ssp'+m.qcod+'.'+m.mmy
 m.dfile = 'd4708'+'.'+m.mmy

 IF !fso.FileExists(lcpath+'\'+m.dfile)
  MESSAGEBOX(CHR(13)+CHR(10)+'���� ����� �� ���������!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(lcpath+'\'+m.ctrl)
  MESSAGEBOX(CHR(13)+CHR(10)+'���� ������ �� �����������!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(lcpath+'\'+m.prsp)
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� �� �����������!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF fso.FileExists(lcpath+'\'+m.ctrln)
  fso.DeleteFile(lcpath+'\'+m.ctrln)
 ENDIF 
 fso.CopyFile(lcpath+'\'+m.ctrl, lcpath+'\'+m.ctrln)
 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 SELECT aisoms
 m.cmessage = ALLTRIM(cmessage)
 USE IN aisoms 

 UnZipOpen(pbase+'\'+gcPeriod+'\'+m.mcod+'\'+m.dfile)
 UnzipGotoFileByName(m.ssp)
 UnzipFile(lcPath+'\')
 UnZipClose()
 IF !fso.FileExists(lcpath+'\'+m.ssp)
  MESSAGEBOX(CHR(13)+CHR(10)+'�� ������� ����������� ����!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 ZipFile = 'rep4708'+m.qcod+'.zip'
 IF fso.FileExists(lcPath+'\'+ZipFile)
  fso.DeleteFile(lcPath+'\'+ZipFile)
 ENDIF 
 
 ZipOpen(lcPath+'\'+ZipFile)
 ZipFile(lcPath+'\'+m.ctrln)
 ZipFile(lcPath+'\'+m.prsp)
 ZipFile(lcPath+'\'+m.ssp)
 ZipClose()
 
 IF !fso.FileExists(lcPath+'\'+ZipFile)
  MESSAGEBOX(CHR(13)+CHR(10)+'���������� ������� �����!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 m.cto      = 'oms@mgf.msk.oms'
 m.cmessage = ALLTRIM(m.cmessage)
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
 poi.WriteLine('Resent-Message-Id: '+m.cmessage)
 poi.WriteLine('Attachment: '+m.dfile+' '+ZipFile)
 
 poi.Close
 
 fso.CopyFile(lcPath+'\'+ZipFile, pAisOms+'\oms\output\'+m.dfile)
 fso.CopyFile(lcPath+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)

 MESSAGEBOX('����� ����������!',0+64,'')

RETURN  