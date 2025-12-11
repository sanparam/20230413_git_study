<%@ Language=VBScript %>
<!--#include file="sql_injection.asp"-->
<%
set cn = Server.CreateObject("ADODB.Connection")
cn.Open Application("db_ConnectionString")


%>
<% 

	flag = session("flag")
	
%>

<html>
<head>
<title>(주)매크로21-MRO전용매장</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="top.css" type="text/css">

<script language="JavaScript">
<!-- 여긴 주석처리부분인가?
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}


function searching(flag) {
    
	if (frm.textfield.value == '' ) {
		alert('검색어를 넣어주세요.');
		frm.textfield.focus();
	}
	
	
	else
	 {
		//frm.next.value = "/ke/search/category_img.asp?method=TopFrame_Search&gubun="+frm.gubun.value +"&Search=" + frm.search.value;
		frm.submit();
	}
}
 
function Login()
{
		window.open("/login/login.asp?path=main", "NewWindow",'Width=520,Height=300,left=20,top=20,toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0');
		
}	
function Login_(target)
{
		target = target.replace("?", "%3F");
		target = target.replace("=", "%3D");
		target = target.replace("&", "%26");
		target = target.replace("?", "%3F");
		target = target.replace("=", "%3D");
		target = target.replace("=", "%3D");
	
		window.open("/login/login.asp?path="+target, "NewWindow",'Width=520,Height=300,toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=no,resizable=0');
}

function Login3()
{	
		window.open("/login/login3.asp?path=main",
		"NewWindow",'Width=520,Height=300,toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=no,resizable=0');
}

function memo(target)
{
	window.open (target, "memo", 'width=700,height=380,left=100,top=100,status=yes,toolbar=no,scrollbars=1,resizable=0');
}

 //-->
</script>



</head>

<body bgcolor="#FFFFFF" text="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" align = "center">
<div class="wrapper" >
  <div class="content">
<% ' top   2002, 2, renewal %>
<table width="880" border="0" cellspacing="0" cellpadding="0" height="80">
<tr>
<td background="images_top/top_bg.gif">
<table width="880" border="0" cellspacing="0" cellpadding="0" height="79">
  <tr>
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="78">
        <tr> 
          <td rowspan="2" width="180" align="center"><a href="index_body.asp" target="content"><img src="images_top/logo.gif" width="162" height="38" border="0"></a></td>
          <td align="right">
            
			
			<table width="255" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><!--home--><a href="index_body.asp" target="content"><img src="images_top/utility_01.gif" width="33" height="13" border="0" alt="MRO 전용매장 메인 페이지로"></a></td>
                 <%if session("id") = "" then%>
				<td><!--log in--><a href="javascript:Login_('/frame.asp?main_url=history/default.asp');"><img src="images_top/utility_02.gif" width="44" height="13" alt="로그인" border="0"></a></td>
                 <%else%>
				<td><a href="/login/logout.asp" target="_top" ><img src="images_top/utility_02_out.gif" width="44" height="13" alt="로그아웃" border="0"></a></td>
				 <%end if%>
				<td><!--회원가입-->
				<%if session("id") = "" then%>
				  <a href="manage/member_inSE_ke.asp" target="content" ><img src="images_top/utility_04.gif" width="54" height="13" alt="Registration" border="0"></a>
				 <% End If %>
				</td>
				<td><!--site map--><a href="sitemap.asp" target="content"><img src="images_top/utility_03.gif" width="48" height="13" alt="사이트맵" border="0"></a></td>
                <td><!--http://www.mangchi.co.kr<a href="/aboutus/index.html" target="content"><img src="images_top/utility_05.gif" width="76" height="13" alt="(주)매크로21 회사소개" border="0"></a>--></td>
              </tr>
            </table>
			
          
		  
		 </td>
        </tr>
        <tr> 
          <td bgcolor="#000000" height="30"> 
            <table width="650" border="0" cellspacing="0" cellpadding="0" class="txt_white">
              <tr valign="middle" align="center"> 
                <td><a href="info.asp"  target="content"><b>통합구매</b>소개</a>&nbsp;&nbsp;&nbsp</td>
                <td><%if session("id") = "" then%>
					 <a href="javascript:Login_('/frame.asp?main_url=history/default.asp');">
					<%else%>
					 <a href="history/default.asp" target="content">
					<%end if%>
					My Page &nbsp;&nbsp;&nbsp;
					</a>
				</td>
<!--                <td><a href="auction/re_good_list.asp" target="content">입찰신청</a></td> -->
                <td><!--a href="mailto:webmaster@macro21.com"--><%if session("id") = "" then%>
		   			<a href="javascript:Login3();">
		  			<%else%>
		  			<a href="offer/write.asp" target="content">
		  			<%end if%>
					&nbsp;문 의&nbsp;&nbsp;&nbsp;
					</a></td>
                <td><a href="commerce/basket_main.asp" target="content">&nbsp;카트(장바구니)</a>&nbsp;&nbsp;&nbsp;</td>
                <td><a href="javascript:window.open('http://192.168.0.11:81')">센서/스위치전용매장</a>&nbsp;&nbsp;</td>
				<td><a href="javascript:window.open('http://192.168.0.11:8089')">사업안내</a></td>
			   </tr>
            </table>
          </td>
        </tr>
        <tr align="center">
		  
		  
		  <!--top- "깜뎅이님 안녕하세요." -->
		  <td colspan="2" class="txt_white"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="180" class="txt_white" align="center">
				<!--쪽지시작-->
                <!--#include file="memo/inc.asp" -->
                <% cnt = new_count(Session("ID")) %>
                <% if  cnt >= 1 then '새로운 쪽지가 있을때만 나온다.%>
					
					<% set new_cnt= server.CreateObject("ADODB.Recordset")
						new_cnt.ActiveConnection = cn
				
						sql = "select top 1 n.*, m.usernm as from_name from note n join muser m on n.[from] = m.customid where [to] = '" & session("id") & "' and [read] = 'N' order by date desc "
						new_cnt.Open sql
					%>
					<a href="javascript:memo('memo/detail.asp?no=<%=new_cnt("no")%>')">새 쪽지<%=cnt%>개 도착&nbsp;<img SRC="images_top/new_ani.gif" border="0" WIDTH="21" HEIGHT="9"></a>
				<% end if %>	
				<!--쪽지 끝-->				  
                <td class="txt_white">&nbsp;
                <%if session("id") <> "" then%>&quot;
		   		  <%=Session("UserName")%>님 안녕하세요.&quot;
		   		<%end if%>
                </td>
              </tr>
            </table>
          </td>
		  
		  
        </tr>
      </table>

    </td>
	
  <!--/tr>
  
  
</table>
</td-->

<!--td width="150"-->
<!--search table-->		
<!--form name="frm" action="../search/category_img.asp" method="POST" target="content">
  <table valign="bottom", align="left" background="images_top/top_bg.gif" width="100%" border="0" cellspacing="0" cellpadding="0" height="79" class="txt-33">
     <tr> 
       <td valign="bottom" align="left" height="53">
	      <select name="keyfield">
          <option value="product" selected>제품명</option>
          <option value="company">제조사</option>
          <option value="pageno">페이지 번호</option>
          </select>
       </td>
	 </tr>
	 <tr>
       <td valign="top" align="left" height="20">
	  <input type="text" name="textfield" maxlength="30" size="15">
	  <a href="javascript:searching('<%=session("flag")%>');">
	  <img src="../../images_home/search_go.gif" alt="키워드를 넣고 검색하세요" border="0" width="25" height="18" align="absmiddle"></a>
       </td>
      </tr>
    </table>
   </form-->
	  <!--search table end-->
<!--/td-->
  <!--td background="images_top/top_bg.gif"><!-- Search Google -->
<!--center>
<form method="get" action="http://www.google.co.kr/custom" target="content">
<table bgcolor="#000000" width ="150">
<tr><td nowrap="nowrap" valign="top" align="left" height="13">
<a href="http://www.google.com/">
<img src="http://www.google.com/logos/Logo_25blk.gif" border="0" alt="Google" align="middle"></img></a>
<input type="text" name="q" size="12" maxlength="255" value=""></input>
<input type="submit" name="sa" value="검색"></input>
<input type="hidden" name="client" value="pub-7139684911602044"></input>
<input type="hidden" name="forid" value="1"></input>
<input type="hidden" name="ie" value="EUC-KR"></input>
<input type="hidden" name="oe" value="EUC-KR"></input>
<input type="hidden" name="cof" value="GALT:#003324;GL:1;DIV:#66CC99;VLC:FF6600;AH:center;BGC:C5DBCF;LBGC:73B59C;ALC:000000;LC:000000;T:330033;GFNT:333300;GIMP:333300;LH:50;LW:50;L:http://www.google.com/images/logo.gif;S:http://;LP:1;FORID:1;"></input>
<input type="hidden" name="hl" value="ko"></input>
</td></tr></table>
</form>
</center>
<!-- Search Google -->

  <!-- /td-->
  </tr>
  <tr>
  	<td height="1"></td>
  </tr>
</table>

<!--top end-->
   </div>
</div>
</body>
</html>
