*----------------------------------
*	��ֹ���򱻶�ο���
*----------------------------------
Public handle
Declare Integer CreateFileMapping In kernel32.Dll Integer hFile, ;
	INTEGER lpFileMappingAttributes,Integer flProtect, ;
	INTEGER dwMaximumSizeHigh, Integer dwMaximumSizeLow, ;
	STRING lpName
Declare Integer GetLastError In kernel32.Dll
Declare Integer CloseHandle In kernel32.Dll Integer hObject

szname="xmd-sos"   
handle = CreateFileMapping(0xFFFFFFFF,0,4,0,128,szname)

If handle = 0
	Wait Window "CreateFileMapping ʧ�� - LastError: " + Ltrim(Str(GetLastError()))
	Return
Endif

If handle=0
	Messagebox("�ڴ�ӳ���ļ�����ʧ�ܣ�",16,"����")
	Return .F.
Else
	If GetLastError()=183
		=Messagebox("�ó�������Ѿ����У�",16,"��ʾ")
		Close All
		Clear Dlls
		Clear Events
		Quit
	Endif
ENDIF

**** ϵͳ��ʼ ----
lcLastSetPath=SET("PATH")
CD "\xmd\"
SET PATH TO;DATA;FORMS;LIBS;MENUS;PROGS;BMP;EXCEL
***
SET SYSMENU TO    && �������������� SET SYSMENU TO �����ֹ Visual FoxPro ���˵���
MODIFY WINDOW SCREEN FROM 0.00, 0.00 TO 48.00,168.50 &&�������� , 1024X768
_SCREEN.Caption = "��¼ϵͳ"
_SCREEN.Icon = "..\bmp\NET01.ICO"
_SCREEN.Picture = "..\bmp\main_bg.gif"
_SCREEN.CLOSABLE = .F.  && ȡ�����ڹرձ���ť

**********************************
DO ..\PROGS\ϵͳ����.prg

SET STATUS BAR ON   && ��ʾ����ȥͼ��״̬�� ,on Ĭ��״̬��
SET SYSMENU TO      && �������������� SET SYSMENU TO �����ֹ Visual FoxPro ���˵���
SET ECHO OFF        && (Ĭ��ֵ)�ر� FoxPro 2.0 ��ǰ�汾�еĸ��ٴ���
SET ESCAPE OFF      && ��ֹ���еĳ���������ڰ� Esc �����жϡ�
SET SAFETY OFF      && ������д�����ļ�֮ǰ�Ƿ���ʾ�Ի���
SET TALK OFF        && ���� Visual FoxPro �Ƿ���ʾ������
SET CENT ON         && ��ʾ���Ϊ4λ
SET DATE ANSI       && yy.mm.dd (�����ո�ʽ)
SET DELETE ON       && ON Ϊʹ�÷�Χ�Ӿ䴦���¼(��������ر��еļ�¼)��������Ա���ɾ����ǵļ�¼��
SET EXCL OFF        && (˽�����ݹ����ڵ�Ĭ�Ϸ�ʽ)���������ϵ��κ��û�������޸������ϴ򿪵ı�
SET HELP ON         && ����򿪰���

* ��������������
_SCREEN.Caption = "��¼ϵͳ"
_SCREEN.Icon = "..\bmp\NET01.ICO" 
_SCREEN.Picture = "..\bmp\main_bg.gif"
MODIFY WINDOW SCREEN FROM 0.00, 0.00 TO 48.00,168.50 &&�������� , 1024X768

* ϵͳ��¼ 
@44,84 SAY c�汾�� ;
 	 FONT "Arial" ,15 ;
     STYLE "BIUT";
     COLOR RGB(128,128,128,,,)   
     
DO FORM '..\forms\login.scx'

IF mem_on="OFF"
* ɾ����ʱ�� 
CLOSE DATABASES ALL
CLOSE TABLES ALL 
SELECT 0
USE ..\data\mmk1.dbf EXCLUSIVE 
ZAP
USE
DO ..\PROGS\quit.prg
ENDIF 

IF  mem_on="ON"
  CLOSE DATABASES ALL
  CLOSE TABLES ALL 
  SELECT 0
  USE ..\data\mmk1.dbf EXCLUSIVE 
  ZAP
  USE

  * �ָ����洰��
  MODIFY WINDOW SCREEN FROM 0.00, 0.00 TO 48.00,168.50 &&�������� , 1024X768
  * �ָ�ϵͳ�˵�
  * SET SYSMENU TO DEFAULT
  DO '..\progs\xmd.PRG'

ELSE
  * �ָ�ϵͳ�˵�
  SET SYSMENU TO DEFAULT
  * �ָ����洰��
  MODIFY WINDOW SCREEN FROM 0.00, 0.00 TO 48.00,168.50 &&�������� , 1024X768
  SET STATUS BAR ON
  CLEAR
  RETURN
 
ENDIF