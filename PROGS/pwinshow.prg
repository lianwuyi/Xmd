RELEASE c����,c��ʾ
PUBLIC  c����,c��ʾ

*** ����ǰ���쵽�����ѡ�
IF c���� - DATE() < 3  && ��ǰ������ʾ���ڡ�
   c���� = "��Wwwjxc �������ѡ�"
   c��ʾ = "�������"+ALLTRIM(STR(c���� - DATE()))+"����ڣ�;
   Ϊ����������ʹ�ã��뼰ʱ�빩Ӧ����ϵ��"+CHR(13)+c����֧��
   DO ..\progs\popwindows.prg
ENDIF

IF c�̵��¼ = "1" && ӵ���̵�Ȩ�޵ģ�����ʾ
  *** ����ǰ28�ž��̵����ѡ�
  IF DATE() = ctod(subs(dtoc(date()),1,8)+'30') OR DATE() = ctod(subs(dtoc(date()),1,8)+'28');
     OR DATE() = ctod(subs(dtoc(date()),1,8)+'29') OR DATE() = ctod(subs(dtoc(date()),1,8)+'31')
     c���� = "��Wwwjxc �̵����ѡ�"
     c��ʾ = "�̵������ѽӽ����������̵�׼����"
     DO ..\progs\popwindows.prg
  ENDIF

  *** ��1���̵����ѡ�
  IF DATE() = ctod(subs(dtoc(date()),1,8)+'1') 
     c���� = "��Wwwjxc �̵����ѡ�"
     c��ʾ = "����Ϊϵͳ�̵����� 1�ţ��뽫���¶�����¼�꣬�°������̵㣡"
     DO ..\progs\popwindows.prg
  ENDIF 
ENDIF 