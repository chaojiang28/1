<%
'****************************************************
' �������ƣ���Դ������ѵ�ӰС͵ ty  1.8��
' ������������Դ����
' ����֧��: http://bbs.88ag.com
' ���ص�ַ: http://bbs.88ag.com     ����ͣ���£�����ʻ�ȡ���³���
'****************************************************
%>

<%
Function GetHttpPage(HttpUrl)
   If IsNull(HttpUrl)=True Or HttpUrl="$False$" Then
      GetHttpPage="$False$"
      Exit Function
   End If
   Dim Http
   Set Http=server.createobject("MSXML2.XMLHTTP")
   Http.open "GET",HttpUrl,False
   Http.Send()
   If Http.Readystate<>4 then
      Set Http=Nothing 
      GetHttpPage="$False$"
      Exit function
   End if
   GetHTTPPage=bytesToBSTR(Http.responseBody,"utf-8")
   Set Http=Nothing
   If Err.number<>0 then
      Err.Clear
   End If
End Function

Function BytesToBstr(Body,Cset)
   Dim Objstream
   Set Objstream = Server.CreateObject("adodb.stream")
   objstream.Type = 1
   objstream.Mode =3
   objstream.Open
   objstream.Write body
   objstream.Position = 0
   objstream.Type = 2
   objstream.Charset = Cset
   BytesToBstr = objstream.ReadText 
   objstream.Close
   set objstream = nothing
End Function

Function GetBody(ConStr,StartStr,OverStr,IncluL,IncluR)
   If ConStr="$False$" or ConStr="" or IsNull(ConStr)=True Or StartStr="" or IsNull(StartStr)=True Or OverStr="" or IsNull(OverStr)=True Then
      GetBody="$False$"
      Exit Function
   End If
   Dim ConStrTemp
   Dim Start,Over
   ConStrTemp=Lcase(ConStr)
   StartStr=Lcase(StartStr)
   OverStr=Lcase(OverStr)
   Start = InStrB(1, ConStrTemp, StartStr, vbBinaryCompare)
   If Start<=0 then
      GetBody="$False$"
      Exit Function
   Else
      If IncluL=False Then
         Start=Start+LenB(StartStr)
      End If
   End If
   Over=InStrB(Start,ConStrTemp,OverStr,vbBinaryCompare)
   If Over<=0 Or Over<=Start then
      GetBody="$False$"
      Exit Function
   Else
      If IncluR=True Then
         Over=Over+LenB(OverStr)
      End If
   End If
   GetBody=MidB(ConStr,Start,Over-Start)
End Function

Function ShowErr(ErrMsg)
Response.write "<script>alert('" & ErrMsg &"');history.back();</script>"
Response.end
End Function
%>








<%
FsoName="Scripting.FileSystemObject"
AdoName="ADODB.Stream"
function checkCache()'����Ƿ����ĳ��cache,����cache����һ��ˢ��ʱ��
if isempty(application(cacheName)) or (not isdate(application(cacheName&"_RefreshTime"))) then
checkCache=false
else
checkCache=true
end if
end function
sub setTimer() '����ˢ��ʱ��Ϊ��ǰʱ��
application.Lock()
application(cacheName&"_RefreshTime")=Now()
application.UnLock()
end sub
sub setCache(value)
    if value<>readCache() And value<>"" then
application.Lock()
application(cacheName)=value
application.UnLock()
end if
end sub
function readCache()
readCache=application(cacheName)
end function
function checkTimer()'����Ƿ��Ѿ�����ˢ�¼�����ǻ���true �񷵻�false
if DateDiff("s", application(cacheName&"_RefreshTime"), now)>=RefreshTime then
checkTimer=true
else
checkTime=false
end if
end function
sub makeLocalFile()'���ɱ����ļ�!!
if useObject="Fso" then
    SaveToFileUseFso readCache(),server.mappath(localsavePath)
elseif useObject="Ado" then
    SaveToFileUseAdo readCache(),server.mappath(localsavePath)
end if
end sub
sub SaveToFileUseFso(ByVal strBody,ByVal File)
On Error Resume Next
set fso=createobject(Fsoname)
if Err<>0 then
   Response.Write "����Fso�������!����Ӷ�һ������ָ�����Fso�������ʵ����:ʵ��.FsoName='....'"
   Err.Clear
   Response.End
end if
set cacheMake=fso.createtextfile(File,true)
cacheMake.writeline strBody
cacheMake.close
set cacheMake=nothing
end sub
sub SaveToFileUseAdo(ByVal strBody,ByVal File)
    Dim objStream
    On Error Resume Next
    Set objStream = Server.CreateObject(AdoName)
    If Err.Number=-2147221005 Then 
        Response.Write "����Ado�������!����Ӷ�һ������ָ�����Ado.stream����ʵ����:ʵ��.AdoName='....'"
        Err.Clear
        Response.End
    End If
    With objStream
        .Type = 2
        .Open
        .Charset = "GB2312"
        .Position = objStream.Size
        .WriteText = strBody
        .SaveToFile File,2
        .Close
    End With
    Set objStream = Nothing
End sub 
sub fresh()
        if checkCache()=false or checkTimer()=true then'���δ������,�����Ѿ���ʱ
           setCache getText()
     setTimer()
     if needLocal=true then'�����Ҫ���������ļ�
           makeLocalFile()
     end if
  end if
  if needDisplay=true then'ˢ��cache���ڱ�ҳ����ʾ
  response.write readCache()'�ڱ�ҳ����ʾ
  elseif goToWeb=true then
  response.redirect localsavePath
        end if
end sub

%>




