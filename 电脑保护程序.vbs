function guanji()
	set ws=createobject("wscript.shell")
	ws.run"cmd.exe /c shutdown -s"
End function
do
dim a
a=inputbox("��������������������")
if a="lmc070815" Then
	msgbox"�������ã�������ʹ�õ���"
	msgbox"���ע���ں�LMCCPAS�����ָ�������˼��С����"
	exit do
else 
	msgbox"���ԣ����Խ���1�����ڹػ�������������Իش������Խ���ػ�"
	dim b	
	b=inputbox("���������⣺������������������죿�������£�xxxxxxxx �ޡ�-�����ޡ� ��")
		if b="20070815" Then 
			dim c
			c=inputbox("Ϊ����֤��ݣ�����һ�⣬����巵�Ӣ������ʲô��������ĸ��д��")
			if c="William"	Then 
				msgbox"���ˣ����������lmc070815Ŷ"
				msgbox"���ע���ں�LMCCPAS�����ָ�������˼��С����"
			exit do
			else guanji()
			msgbox"�����ػ�"
			end If
		else	
			msgbox"�����ػ�"
			msgbox"���ע���ں�LMCCPAS�����ָ�������˼��С����"
			guanji()
		end If
end If
loop
