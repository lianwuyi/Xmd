* �����﷨��
* DoDefault() && ����ִ��Ĭ�ϴ���  
* thisform.��Ʒ����1.setfocus  &&���

* && ȡ����   floor(val(str(n_cfid)))       floor() ȡ����ȥ���ո�,  int() ȡ���ǲ�ȥ���ո� CEILING(3.1) = 4   int(3.1) = 3  

* ThisForm.Refresh()  && ˢ�±�
* THISFORM.Grid1.Refresh()  && ˢ�±��
* ��Ʒ����: д��2 = 'Y'  rele ���s,��Ʒ����s,Ʒ�����s,��λs,����s,���ܺ�s,��עs,��ɫs

*��ӡ����REPORT FORM c:\jck\forms\�ƻ���Ʒ��.frx NOEJECT NOCONSOLE TO PRINTER   
*        REPORT FORM ..\forms\dzbg.frx NOEJECT NOCONSOLE PREV
******************************************************
** ���水ť
*  =cursorsetprop("Buffering",5)
*   =tableupdate(.t.)             && ��������
*  thisform.cmd���.enabled=.t.
*  thisform.cmd����.enabled=.f.
*  thisform.cmd����.enabled=.f.
*  thisform.cmd�޸�.enabled=.t.
*  ThisForm.Refresh()

*  =cursorsetprop("Buffering",2)
*  wait window '�ѱ�������ʱ�������' nowait noclear
******************************************************

* 1024 * 768  ������ 630 ��880

*   TOTAL TO TableName ON FieldName   [FIELDS FieldNameList]   [Scope]
*       [FOR lExpression1]   [WHILE lExpression2]   [NOOPTIMIZE]

****************************************************************
* IF file("..\ck\arjs\&dbfs"+".arj")=.T.

*  wait windows "�ļ�����,���������룡" AT 8,30 TIMEOUT  2
* return
* ENDIF
** -g:10748612 Ϊ������ѹ����-g: ���뿪�أ� "10748612" Ϊ����
**��ѹʱ��Ҫ�������뿪�ز��ܽ�ѹ��


* run arj a -g:10748612 ..\ck\arjs\&dbfs  ..\ck\data\*.*
* run arj e -g:10748612  "\fapiao\arjs\&dbfs_1

********************************************************
*     INDEX ON eExpression TO IDXFileName | TAG TagName [OF CDXFileName]
*     [FOR lExpression]
 *    [COMPACT]
  *   [ASCENDING | DESCENDING]
   *  [UNIQUE | CANDIDATE]
    * [ADDITIVE]
****

 *************************
*   select *  from cPATHS+'syslydj.dbf' ;
*        where alltrim(��ҵ������) == djh1 ;
*        into table ..\test.dbf
*      into cursor tmp

*  UPDATE cPATHS+'syslydj.dbf' SET ��ɼǺ� = "*" , ������� = rq1 WHERE alltrim(��ҵ������) == djh1
 **************************        
 
 ***���� Excel ��************************************
*  USE c:\bjxt\data\�����.dbf AGAIN IN 0 ALIAS �����
*    SELECT �����
*    BROWSE LAST
*    COPY TO "c:\bjxt\e-mail_html\2.xls" TYPE XL5
*
********************************************************
*��ӡ����REPORT FORM c:\jck\forms\�ƻ���Ʒ��.frx NOEJECT NOCONSOLE TO PRINTER  


*�³�������

***  SQLDISCONNECT(0)  && �Ͽ��������ݿ�����

*** SQL �� VFP �������ֶδ���
* 1�� SQL �� VFP ���е������ֶζ��б���Ϊ���������ֵ .NULL.
* 2)  ����һ��������ֵ�� STORE .NULL. TO rq1
*    ���������� repl �������� with rq1 for �������� = {    .  .  }
*    ������VFP ��������Ϊ�յľͿ��Ը����ֵ�� .NULL.
** 
** ��Ʊ�ʱ���������ֵ����ֵ���ֶ�Ҫ�����ֵ0�������ڸ�����Ӽ�¼ʱ���粻�Կ�ֵ����¼����ֵ��0 �ͻ������ INSERT INTO ������ ��
** ��Ʊ�ʱ���Թؼ��ֶ�Ҫ������������������߲�ѯ���޸ġ������¼ʱ�ſ졣


***************************************************************
* ��SQL2000��ҵ����������������NULL�ַ���
* ctrl+0,����������<null>��䣬���Ѿ���<null>��ֵ���˸��ֶΣ�����ȥ������������������ֶ��ϵ������ῴ��<null>�Ѿ�д��ȥ��
* SET NULLDISPLAY TO '' && ȥ��.NULL
* WAIT WINDOW NOWAIT "���ݿ��¼д���: ��¼" + ALLTRIM(STR(kk)) + " of " + ALLTRIM(STR(ss))
* set message to  && ���״̬����Ϣ