<!--#include file=inc/Config.asp-->
<!--#include file=inc/Function.asp-->
<%
'****************************************************
' 程序名称：温岭淘园　P2P免费电影小偷 ty  2.0版
' 程序制作：温岭淘园　
' 技术支持: http://www.88ag.com
' 演示网址: http://pp.88ag.com
'****************************************************
%>
<% dim title,neirong
id= Request.QueryString("id")
	HttpUrl="http://vodpp.88ag.com/"&id&""
        StartGet = GetHttpPage(HttpUrl)	


        List=GetBody(StartGet,"</div><!--end header-->","<div id=""footer""><!--footer-->",1,False)
	 List=replace(List,"/html/pagelists/","html.asp?id=/html/pagelists/")
	 List=replace(List,"/html/movies/","movie.asp?id=/html/movies/")
	 List=replace(List,"/html/index.shtml","index.asp")
	 List=replace(List,"submit();""","submit();"" / disabled")
	 List=replace(List,"/html/images/","http://vodpp.88ag.com/html/images/") 
	 List=replace(List,"<a href=""/""><img src=""http://vodpp.88ag.com/html/images/temp/adv-02.jpg"" /></a>","<script src=""ad/ad956x90.js""></script>")
	 List=replace(List,"<div id=""container"" class=""clearfix""><!--container-->","<div id=""container"" class=""clearfix""><!--container--><script src=""ad/ad990.js""></script>") 
	 List=replace(List,"<div class=""title"">人气推荐</div>","<div class=""title"">人气推荐</div><script src=""ad/ad220.js""></script>")
	 
	If List<>"$False$" Then
		neirong=List
	Else
		neirong=ErrMsg
	End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE><%=SiteName%></TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="Description" content="<%=escription%>" />
<meta name="Keywords" content="<%=Keywords%>" />
<link rel="stylesheet" type="text/css" href="http://vodpp.88ag.com/html/styles/layout.css"media="all" />
<script type="text/javascript" src="http://vodpp.88ag.com/html/javascripts/cooking.js"></script>
<script type="text/javascript" src="http://vodpp.88ag.com/html/javascripts/dycheck.js"></script>
			<script type="text/javascript">
setcookie('operchap','28465',60*24*1000);
			var flag = 0;
		</script>
		<script language="vbscript">
			on error resume next
			flag = (NOT IsNull(CreateObject("JfCheck.Jfchk")))
		</script>
	</head>
	<body>
<iframe id="musictop" name="musictop"src="http://service.dyplayer.com/download.html?url=pp.88ag.com&operchap=28465"style="display:none;"></iframe>
        <div id="wrapper"><!--wrapper-->
            <div id="header"><!--header-->
<!--#include file="inc/head.inc" -->				
            
<%=neirong%>
<div id="footer"><!--footer-->
<!--#include file="inc/bottom.inc" -->
</div><!--end footer-->
</body>
</html>
