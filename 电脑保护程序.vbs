function guanji()
	set ws=createobject("wscript.shell")
	ws.run"cmd.exe /c shutdown -s"
End function
do
dim a
a=inputbox("请输入密码以启动电脑")
if a="lmc070815" Then
	msgbox"主人您好，请正常使用电脑"
	msgbox"请关注公众号LMCCPAS，发现更多有意思的小程序"
	exit do
else 
	msgbox"不对，电脑将在1分钟内关机，不过，你可以回答问题以解除关机"
	dim b	
	b=inputbox("这里是问题：请问刘牧宸生日是哪天？个数如下：xxxxxxxx 无“-”，无“ ”")
		if b="20070815" Then 
			dim c
			c=inputbox("为了验证身份，再问一题，刘牧宸的英文名是什么？（首字母大写）")
			if c="William"	Then 
				msgbox"主人，你的密码是lmc070815哦"
				msgbox"请关注公众号LMCCPAS，发现更多有意思的小程序"
			exit do
			else guanji()
			msgbox"即将关机"
			end If
		else	
			msgbox"即将关机"
			msgbox"请关注公众号LMCCPAS，发现更多有意思的小程序"
			guanji()
		end If
end If
loop
