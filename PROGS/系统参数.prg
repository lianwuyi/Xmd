
SET SAFETY OFF    && ������д�����ļ�֮ǰ�Ƿ���ʾ�Ի���
SET TALK OFF      && ���� Visual FoxPro �Ƿ���ʾ������
SET CENT ON       && ��ʾ���Ϊ4λ
SET DATE ANSI     && yy.mm.dd (�����ո�ʽ)
SET DELETE ON     && ON Ϊʹ�÷�Χ�Ӿ䴦���¼(��������ر��еļ�¼)��������Ա���ɾ����ǵļ�¼��
SET EXCL OFF      && (˽�����ݹ����ڵ�Ĭ�Ϸ�ʽ)���������ϵ��κ��û�������޸������ϴ򿪵ı�
SET HELP ON       && ����򿪰���
SET MULTILOCKS ON && ������

* SET DEFAULT TO c:\
* ���ó�����Ŀ¼
RELEASE gcMainPath
PUBLIC gcMainPath

gcMainPath = Sys(5)+"\" 
Set Default To &gcMainPath 

* gcMainPath = "c:\"
CD "\xmd\"
SET PATH TO;DATA;FORMS;LIBS;MENUS;PROGS;BMP;EXCEL

RELEASE cPATHS,c�汾��,c��˾,c�������,c�绰,c��ַ,c����֧��,c����
PUBLIC cPATHS,c�汾��,c��˾,c�������,c�绰,c��ַ,c����֧��,c����
SELECT 0
USE ..\DATA\sys1.DBF
cPATHS = ALLTRIM(��������)
c��˾ = ALLTRIM(��˾��)
c�绰 = ALLTRIM(�绰)
c��ַ = ALLTRIM(��ַ)
c������� = 1
USE
c�汾�� = 'ϸ�뵥����ϵͳ 1.0 180808 ��'
c����֧�� = '����֧��:  lianwuyi@163.com '
c���� = {^2025.01.01}

*** ----------------------------------------------------
* ��ֹ�ͻ�����ʹ�����
IF DATE() > c����
  MESSAGEBOX('����������ѵ�,���빩Ӧ�̹���ʹ��Ȩ!'+CHR(13)+CHR(13)+c����֧��;
              +chr(13),16,c�汾��)
  WAIT CLEAR
  CLOSE All
  CLEAR DLLS
  CLEAR Events
  QUIT
  RETURN
ENDIF
*** -----------------------------------------------------

*** �������ݿ� ***
IF FILE(cPATHS+'mmk.dbf') = .T.
SET REPROCESS TO 30  && ���������Ĵ���Ϊ 30�� 
SET EXCLUSIVE ON && OFF 
SET DELETED ON

  SELECT 0
  USE cPATHS+'mmk.dbf'
  IF FLOCK()
    *WAIT WINDOW "��ʾ�������������ݿ�..." NOWAIT TIMEOUT 3
    COPY all to ..\test.dbf 
    USE 
  
    SELECT 1
    USE ..\data\mmk1.dbf 
    ZAP
    APPEND FROM ..\test.dbf
    USE 
    DELETE FILE ..\test.dbf 
    *WAIT WINDOW "��ʾ����ѯ�ɹ���" NOWAIT TIMEOUT 3
  ELSE 
    USE 
    WAIT WINDOW "�����������ݿ�ʧ�ܣ����Ժ����ԡ���" TIMEOUT 4
    QUIT 
    RETURN
  ENDIF 
  
ELSE 
  WAIT WINDOW "�����Ҳ������ݿ��ļ������Ժ����ԡ���" TIMEOUT 4
  QUIT 
  RETURN
ENDIF