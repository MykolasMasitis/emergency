PROCEDURE m_menu
SET SYSMENU TO

DEFINE PAD mmenu_1 OF _MSYSMENU PROMPT '\<���������� �� ���' COLOR SCHEME 3 KEY ALT+A, ""
DEFINE PAD mmenu_5 OF _MSYSMENU PROMPT '\<������' COLOR SCHEME 3 KEY ALT+F , ""
ON PAD mmenu_1 OF _MSYSMENU ACTIVATE POPUP popInfFrLpu
ON PAD mmenu_5 OF _MSYSMENU ACTIVATE POPUP popTuneUp

DEFINE POPUP popInfFrLpu MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popInfFrLpu PROMPT '����� ��������� �� ��� (��� ���)'
DEFINE BAR 02 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 03 OF popInfFrLpu PROMPT '��������� �������������� �������' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 04 OF popInfFrLpu PROMPT '��������� �������������� �������' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 05 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 06 OF popInfFrLpu PROMPT '������������ ����� ������ ��� ���'
DEFINE BAR 07 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 08 OF popInfFrLpu PROMPT '��������� ����� � ������'
DEFINE BAR 09 OF popInfFrLpu PROMPT '�������������� ������ � LPU2SMO'
DEFINE BAR 10 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 11 OF popInfFrLpu PROMPT '�����'

ON SELECTION BAR 01 OF popInfFrLpu DO FORM MailView
ON SELECTION BAR 03 OF popInfFrLpu DO FORM MailTrash
ON SELECTION BAR 04 OF popInfFrLpu DO FORM MailDouble
ON SELECTION BAR 06 OF popInfFrLpu DO MakeCTRLs
ON SELECTION BAR 08 OF popInfFrLpu DO Snd2MgfOMS
ON SELECTION BAR 09 OF popInfFrLpu DO Exp2Lpu
ON SELECTION BAR 11 OF popInfFrLpu CLEAR EVENTS 

DEFINE POPUP popTuneUp MARGIN RELATIVE SHADOW COLOR SCHEME 4
DEFINE BAR 01 OF popTuneUp PROMPT '����� ��������� �������' 
DEFINE BAR 02 OF popTuneUp PROMPT '��������� ������� ����������'
DEFINE BAR 03 OF popTuneUp PROMPT '��������� ������' 
DEFINE BAR 04 OF popTuneUp PROMPT '\-'
DEFINE BAR 05 OF popTuneUp PROMPT '�������������� �� ���'
DEFINE BAR 06 OF popTuneUp PROMPT '�������������� ������� ���'
DEFINE BAR 07 OF popTuneUp PROMPT '\-'
DEFINE BAR 08 OF popTuneUp PROMPT '��������� ����� ������' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 09 OF popTuneUp PROMPT '��������� ������� ����'  SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 10 OF popTuneUp PROMPT '\-'
DEFINE BAR 11 OF popTuneUp PROMPT '�������� ����� ������' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 12 OF popTuneUp PROMPT '������� CTRL-��' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 13 OF popTuneUp PROMPT '������� ����� ��������' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 14 OF popTuneUp PROMPT '������� BAK-�����' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 15 OF popTuneUp PROMPT '\-'
DEFINE BAR 16 OF popTuneUp PROMPT '������������� ��������� ������� ��' SKIP FOR !INLIST(gcUser,'OMS','USR')

ON SELECTION BAR 01 OF popTuneUp DO FORM SetPeriod
ON SELECTION BAR 02 OF popTuneUp DO FORM TuneBase
ON SELECTION BAR 03 OF popTuneUp goApp.doForm('set_print','settings')
ON SELECTION BAR 05 OF popTuneUp DO ComReind
ON SELECTION BAR 06 OF popTuneUp DO BasReind
ON SELECTION BAR 08 OF popTuneUp DO PackBD
ON SELECTION BAR 09 OF popTuneUp DO PackBDSv
ON SELECTION BAR 11 OF popTuneUp DO ZapEFiles
ON SELECTION BAR 12 OF popTuneUp DO DelAllCtrl
ON SELECTION BAR 13 OF popTuneUp DO DelAllZapros
ON SELECTION BAR 11 OF popTuneUp DO DelBakFiles
ON SELECTION BAR 16 OF popTuneUp DO CorStruct

SET SYSMENU AUTOMATIC
SET SYSMENU ON