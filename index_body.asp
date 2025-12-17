<%@ Language=VBScript %>
<!--#include file="sql_injection.asp"-->
<%

set cn = Server.CreateObject("ADODB.Connection")
	cn.Open Application("db_ConnectionString")

	set iventgood_rs_new = server.CreateObject("ADODB.Recordset")
	iventgood_rs_new.ActiveConnection = cn
 
	
	      sql = "select top 4 c.catalogno, c.catalognm, c.maker, i.path " &_
	      "from catalog c join dicatalog i on c.catalogno = i.catalogno and i.no='01' " &_
	      "where c.catalogno > 7449 order by c.catalogno"
	iventgood_rs_new.Open sql
	
	'new_type = true 
		
	set iventgood_rs_new_add = server.CreateObject("ADODB.Recordset")
	iventgood_rs_new_add.ActiveConnection = cn
 
		sql = "select top 5 c.catalogno, c.catalognm, c.maker, i.path " &_
	      "from catalog c join dicatalog i on c.catalogno = i.catalogno and i.no='01' " &_
	      "join new_add k on c.catalogno = k.catalogno order by k.regdate desc"
	      
	iventgood_rs_new_add.Open sql
	
	set iventgood_rs_insert = server.CreateObject("ADODB.Recordset")
	iventgood_rs_insert.ActiveConnection = cn
	
		sql = "select top 5 c.catalogno, c.catalognm, c.maker, i.path " &_
	      "from catalog c join dicatalog i on c.catalogno = i.catalogno and i.no='01' " &_
	      "join kepco_new k on c.catalogno = k.catalogno" 
 	
	iventgood_rs_insert.Open sql
	
	set product_rs = Server.CreateObject("ADODB.Recordset")
		product_rs.ActiveConnection = cn
		
		sql = "select top 4 c.catalogno, c.catalognm, c.maker, i.path " &_
			  "from catalog c join dicatalog i on c.catalogno = i.catalogno and i.no='01' " &_
			  "join kepco_good k on c.catalogno = k.catalogno " &_
			  "order by k.cnt desc"
	  
	product_rs.Open sql
	
	m = cstr(cint(mid(now(),6,2))-1)
	
	if len(m) = 1 then
		m= "0" & m
	else 
		m = m
	end if
		
	set good = server.CreateObject("adodb.recordset")
	good.ActiveConnection = cn
	
		sql = "select top 5 s.userid, sum(totalamt) as total, m.usernm from macro21.mrequest s "&_
		  "join muser m on s.userid = m.userid where m.customid in (select customid from kepco_user) "&_
          "and m.userid <> 'C000133' and req_date BETWEEN '2003" & m & "01' AND '2003" & m & "30' "&_
          "group by s.userid, m.usernm order by total desc"
    
    good.Open sql
    
    set corp_class_rs = server.CreateObject ("adodb.recordset")
		corp_class_rs.ActiveConnection = cn
		
		'sql ="select * from corp_class where parent_code is null and code <> 3 "	
		sql ="select * from corp_class where parent_code is null order by code "		
	corp_class_rs.Open sql
		
	''// 기획상품...//''	
	set kepco_plan = server.CreateObject ("adodb.recordset")
		kepco_plan.ActiveConnection = cn
		
		sql = "select top 5 c.catalogno, c.catalognm, c.maker, i.path " &_
			  "from catalog c join dicatalog i on c.catalogno = i.catalogno and i.no='01' " &_
			  "join kepco_plan k on c.catalogno = k.catalogno" 
	
	kepco_plan.Open sql	

	
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="text.css" type="text/css">
<link rel="stylesheet" href="top.css" type="text/css">
<script type="text/javascript">

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-31024572-1']);
  _gaq.push(['_setDomainName', 'macro21.com']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>
<!-- 기획상품 시작 -->

<style type="text/css"> 
<!-- 
#floater {position:absolute; visibility:visible} 
-->
<!-- 
#scrollmenu {position:absolute; visibility:visible} 
--> 
</style>
<script language="JavaScript">
<!--
self.onError=null;
currentX = currentY = 0; 
whichIt = null; 
lastScrollX = 0; lastScrollY = 0;
NS = (document.layers) ? 1 : 0;
IE = (document.all) ? 1: 0;
<!-- STALKER CODE -->
function heartBeat() {
if(IE) { 
diffY = document.body.scrollTop; 
//diffY = 0;
diffX = 0; 
}
if(NS) { diffY = self.pageYOffset; diffX = self.pageXOffset; }
if(diffY != lastScrollY) {
percent = .1 * (diffY - lastScrollY);
if(percent > 0) percent = Math.ceil(percent);
else percent = Math.floor(percent);
if(IE) document.all.floater.style.pixelTop += percent;
if(IE) document.all.scrollmenu.style.pixelTop += percent;
if(NS) document.floater.top += percent; 
lastScrollY = lastScrollY + percent;
}
if(diffX != lastScrollX) {
percent = .1 * (diffX - lastScrollX);
if(percent > 0) percent = Math.ceil(percent);
else percent = Math.floor(percent);
if(IE) document.all.floater.style.pixelLeft += percent;
if(IE) document.all.scrollmenu.style.pixelLeft += percent;
if(NS) document.floater.top += percent;
lastScrollY = lastScrollY + percent;
} 
} 
if(NS || IE) action = window.setInterval("heartBeat()",1);
//-->
</script>

<!-- 기획상품 끝 -->

<script language="JavaScript">
<!--
function show(iObject,bool)
{ 
  if (bool == true)
   iObject.style.display="";  
  else
   iObject.style.display="none";  
}





//-->
function check()
{ 
	alert("MRO 매장 가입 지사/지점만이 사용하실 수 있습니다. 로그인한 후 이용하세요");
	//document.location.href="check.asp";
	
}
function Login2()
{
	window.open("/login/login2.asp?path=/history/default.asp", "NewWindow",'Width=520,Height=300,toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=no,resizable=0');
}

function Login(target)
{
	target = target.replace("?", "%3F");
	target = target.replace("=", "%3D");
	target = target.replace("&", "%26");
	target = target.replace("?", "%3F");
	target = target.replace("=", "%3D");
	target = target.replace("=", "%3D");
	
	//alert(target);
	
	window.open("/login/login3.asp?path="+target, "NewWindow",'Width=520,Height=300,toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=no,resizable=0');
	
}

function gonggi(iventnewsno)
{
   win=window.open ('/announce/contents.asp?no='+iventnewsno,'loginwin','width=485,height=320,left=100,top=100,status=yes,toolbar=no,scrollbars=1,resizable=0');
}
  
function announclist(iventnewsno)
{
   win=window.open ('/announce/list.asp?mode=chkkepco&iventnewsno='+iventnewsno,'announcwin','width=485,height=320,left=100,top=100,status=yes,toolbar=yes,scrollbars=1,resizable=0');
}

function js_image_view(image_file) {
	window.open('/search/bigimage.asp?arg='+image_file, 'message', 'width=500, height=500, scrollbars=1, resizable=1');
}

function memo(target)
{
	window.open (target, "memo", 'width=700,height=380,left=100,top=100,status=yes,toolbar=no,scrollbars=0,resizable=0');
}

function webphone(target)
{
	window.open (target, "webphone", 'width=500,height=450,left=100,top=100,status=yes,toolbar=no,scrollbars=1,resizable=0');
}

// 엑셀카탈로그 팝업창 // 
function xls_catalog()
{
		window.open("/folder/pds/catalog/catalog.asp?path=main", "xls_catalog",'Width=370,Height=170,left=20,top=20,toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0');
		
}

// 추석특집 팝업창 // 

function a()
{
	window.open("../baroncnc/default.htm", "NewWindow",'Width=800,Height=800,toolbar=1,location=1,directories=1,status=1,menubar=1,scrollbars=yes,resizable=1');
}

function chsuk_catalog()
{
		window.open("/folder/pds/chsuk/chsuk.asp", "chsuk_catalog",'toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=0,scrollbars=1,left=20,top=20');
		
}


function searching(flag) {
    
	if (frm.textfield.value == '' ) {
		alert('검색어를 넣어주세요.');
		frm.textfield.focus();
	}
	
	else
	 {
		//frm.next.value = "/ke/search/category.asp?method=TopFrame_Search&gubun="+frm.gubun.value +"&Search=" + frm.search.value;
		frm.submit();
	}
}
 
</script>

<link rel="stylesheet" href="text_old.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" onLoad="MM_preloadImages('images_home/button_go_ov.gif')" onLoad="MM_preloadImages('images_home/icon_buy.gif','images_home/icon_tag.gif','images_home/icon_point.gif')" align = "center">
<!-- 기획상품 시작 -->
<div class="wrapper" >
  <div class="content">
<div id=floater style="left:760px; top:1px; width:120px; height:383px; z-index:10">
  <table border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFF00" bgcolor="#FFFFFF" width="121">
    <% do while not kepco_plan.EOF %>
    <tr> 
      <td height="105"> 
        <div align="center"> 
          <table width="90%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td> 
                <div align="center">
				
						<a href="catalogno.asp?catalogno=<% =kepco_plan("catalogno") %>" class=txt-33>
				
			
						<img src="<% =kepco_plan("path") %>" width="90" height="90" border="0">
				</div>
              </td>
            </tr>
            <tr> 
              <td> 
                <div align="center">
                  <p><font class=normal color=#595959>
				
						<a href="catalogno.asp?catalogno=<% =kepco_plan("catalogno") %>" class=txt-33>
				
			
						<% =kepco_plan("catalognm") %>
						<br>
                    </a></font></p>
                </div>
              </td>
            </tr>
            <tr> 
              <td> 
                <div align="center">
                  <p><font class=txt-66>
						<b>
							<% if kepco_plan("maker") <> "" or isnull(kepco_plan("maker")) then  '제조사가 없는경우 공장그림 안나옴~%>
								<img src="images_home/icon_factory.gif" width="19" height="15" alt="이미지 없음">&nbsp;<% =kepco_plan("maker") %>
							<% end if %>
						</b>
					</font></p>
                </div>
              </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
    <% kepco_plan.MoveNext 
    response.flush
	Loop
	%>
        
  </table>
  </div>
<!-- 기획상품 끝 -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td>
<table width="750" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
      <table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="170" valign="top" background="images_home/left_bg.gif" border="0">
            <table width="170" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
				
	  <!--search table-->		
	  <form name="frm" action="../search/category.asp" method="POST" target="content">
	  <table width="160" border="0" cellspacing="0" cellpadding="0" height="131">
        <tr> 
          <td bgcolor="cccccc" valign="top">
		  	<!--검색, 반복구매, 쪽지, 포인트-->
			<!--#include file="include/search.asp"-->
		  </td>
        </tr>
        <tr>
          <td height="9"><img src="images_home/search_bottom.gif" width="170" height="9"></td>
        </tr>
      </table>
	  </form>
	  <!--search table end-->
				
				
				
				
				</td>
              </tr>
              <tr>
                <td>				
				

<!--left 2 table-->
<table width="170" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
	  <!--left 1-->
	  <table width="170" valign="top" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <!--left menu //  Categories Title Images-->
		  <td valign="top"><img src="images_home/left_menu_01.gif" width="170" height="25"></td>
        </tr>
        <tr> 
          <td align="center"> 
            
			
			<!--Categories // menu 내용-->
			<table width="90%" border="0" cellspacing="0" cellpadding="0" class="txt-33">
				<% do while not corp_class_rs.EOF %>
			<tr>
			  <td height="20" onMouseOver="show(document.all.content1<%=corp_class_rs("code")%>,true)" onMouseOut="show(document.all.content1<%=corp_class_rs("code")%>,false)">
			  &nbsp;&nbsp;&nbsp;<img src="images_home/left_dot.gif" width="6" height="6">&nbsp;&nbsp;
				<%
				Response.Write "<font color='#003399'>"
				
					Response.Write "<a onMouseOver='show(document.all.content1" & corp_class_rs("code") & ",true)'  onMouseOut='show(document.all.content1" & corp_class_rs("code") & ",false)' href='search/category.asp?code="& corp_class_rs("code")& "'>"
								
				Response.Write  corp_class_rs("name") & "</font></a> "
				Response.Write "<div id='content1" & corp_class_rs("code") & "' style='display:none; position:absolute; z-index:200; left:javascript:event.i+50; top:javascript:event.i+50' bgcolor='#ffffff' align='center'>"
				
				
				set corp_class_rs1 = server.CreateObject ("adodb.recordset")
				corp_class_rs1.ActiveConnection = cn
				sql ="select * from corp_class where parent_code='" & corp_class_rs("code") &"' order by code"
				
				corp_class_rs1.Open sql
				
				
				'하위분류가 없으면..
				if corp_class_rs1.EOF then
					Response.Write "<table>"
				else 
					Response.Write "<table width='190' height='30' border='0' bgcolor='#999999' cellspacing='0' cellpadding='3'>"
				end if 
				
				Response.Write "<tr><td>"
				Response.Write "<table width='190' height='30' border='0' bgcolor='#999999' cellspacing='1' cellpadding='4'>"
				
				
				
				if not corp_class_rs1.EOF then
					do while not corp_class_rs1.EOF 
					Response.Write "<tr bgcolor=#ffffff>"
					Response.Write "<td>"
					Response.Write "<a  href='search/category.asp?code="& corp_class_rs1("code") & "'>"
					Response.Write  corp_class_rs1("name") & "</a></td></tr>"
					corp_class_rs1.MoveNext
					 'response.flush
					loop
				end if
				set corp_class_rs1 = nothing
				
				Response.Write "</td></tr></table>"
				Response.Write "</table></div>"
				%>
				</td>
			  </tr>

				<% corp_class_rs.MoveNext()
				 'response.flush
				loop 
				
				set corp_class_rs = nothing	
				%>		
            </table>
			<!--Categories // menu 내용 끝-->
			
			
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <!--가입회원사 // title image-->
		  <td height="30"><img src="images_home/left_menu_02.gif" width="170" height="25"></td>
        </tr>
        <tr> 
          <td align="center"> 
            <!--가입회원사 // 컨텐츠 내용-->
            <%sql = "select customnm from kepco_user k join custom c on k.customid = c.customid " & _
					"where k.customid <> 'nowfly' and k.approval = 'Y' order by customnm "    
					set member = Server.CreateObject("ADODB.Recordset")
						member.ActiveConnection = cn
						member.CursorType = 1
						member.Open sql
			%>
			<%'sql = "select customnm from kepco_user k join macro..custom c on k.customid = c.customid " & _
					'"where k.customid <> 'nowfly' and k.approval = 'Y' order by customnm "    
					'set member = Server.CreateObject("ADODB.Recordset")
						'member.ActiveConnection = cn
						'member.CursorType = 1
						'member.Open sql
			%>
			<table width="90%" border="0" cellspacing="0" cellpadding="0" class="txt-33">
              <tr> 
                <td colspan="2">현재 가입 회원사 :</td>
              </tr>
              <tr> 
                <td align="center">총 <font size="4"><Big><b><%=member.RecordCount%></b></big></font>&nbsp;개사</td>
                <td height="42">&nbsp;
				<%'if session("id") = "" then%>
	                <!--a href="javascript:Login('/frame.asp?main_url=member.asp');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('go','','images_home/button_go_ov.gif',1)"-->
				<%'else%>
					<!--a href="member.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('go','','images_home/button_go_ov.gif',1)"-->
				<%'end if%>
                <!--img src="images_home/button_go.gif" width="42" height="42" name="go" border="0"></a-->
                </td>
              </tr>
            </table>
			<!--가입회원사 // 컨텐츠 내용 끝-->
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
	
	<!--이달의 우수회원사 // title image -->
        <tr> 
        <td align="center"> 
          <table width="170" border="0" cellspacing="0" cellpadding="0">
            <tr>
	      <!--엑셀카탈로그-->
	                          <td align=left><a href="javascript:xls_catalog();"><img src="images_home/phone_img.gif" width="160" height="110" border="0"></a></td>
            </tr><br>
	  </table>			          
        </td>
        </tr>
      </table>
      <!--left menu //  Categories, 가입회원사, 이달의 우수 회원사 end-->
	  <!--left 1 end-->
	</td>
  </tr> 
  <tr>
    <td valign="bottom">
	  <!--left 2-->
	                <table width="170" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <!--자료실-->
                        <td align=left><a href="/pds/list.asp"><img src="images_home/data_img.gif" width="160" height="80" border="0" ></a></td>
                      </tr>
                      <tr>
                         <td align=left><a href="javascript:a();"><img src="folder/pds/chsuk/pcnc.gif" width="160" height="90" border="0"></a></td>
                      </tr>
					  <tr>
                         <td align=left><a href="javascript:chsuk_catalog();"><img src="folder/pds/chsuk/product.gif" width="160" height="60" border="0"></a></td>
                      </tr>
                      <!--tr> 
                        <!--공정거래위원회-->
                        <!--td align=left><a href="http://ftc.go.kr/info/bizinfo/communicationList.jsp"><img src="FTC_LINK.gif" width="160" height="80" border="0" ></a></td>
                      </tr-->
					  <!--tr> 
                        <!--공급사 로그인-->
                        <!--td align=left> 
                          <%if session("code") = "" then%><a href="javascript:Login2();">
                          <img src="images_home/login02.gif" width="160" height="70" border="0" alt="공급사 로그인">
                          <%else%><a href="/login/logout2.asp">
                          <img src="images_home/login03.gif" width="160" height="70" border="0" alt="공급사 로그아웃">
                          <%end if%>
                        </td>
                      </tr-->
                     
					 </table>
	  <!--left 2 end-->
	</td>
  </tr>
</table>			
<!--left 2 table end-->		
				
								
				</td>
              </tr>
            </table>
          </td>
          <td width="580" valign="top">
            <table width="580" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
	  
	  
	  
<!--main topimg-->
	  <table width="580" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images_home/img_main01.gif" width="580" height="9"></td>
        </tr>
        <tr> 
          <td> 
            <table width="580" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><img src="images_home/img_main02.gif" width="9" height="121"></td>
                <td bgcolor="0058cc">
                  <table width="571" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="253"><img src="images_home/img_main0303.jpg" width="253" height="121"></td>
                      <td>
<!--공지사항 : 보여지는 글자의 수는 [공지]포함, 띄어쓰기 포함하여 한글 25자로 제한합니당  						잊지마시고...또한, 공지는 3개만 보여집니당.......알'쪄!!!수거!-->
						<table width="100%" border="0" cellspacing="0" cellpadding="0" class="txt_white">
                          <tr> 
                            <!--td> 
                              <!--곧바로 올라갑니다...-->
                              <!--object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="318" height="50">
                                <param name=movie value="images_home/top_text.swf">
                                <param name=quality value=high>
								<param name="wmode" value="transparent">
								<param name="play" value="loop">
                                <embed src="images_home/top_text.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="318" height="50">
                                </embed> 
                              </object></td-->
							  <td> <img src="images_home/img_title.gif" width="318" height="50">
                             </td>
                          </tr>
                          <%
							Set RS=Server.CreateObject("ADODB.Recordset") 
							rs.ActiveConnection = cn
							Sql = "SELECT TOP 2 * FROM macro21.notice ORDER BY write_date DESC"
							rs.Open sql
						  %>
                          <tr> 
                            <td align="center" height="71" valign="top"> 
                              <table width="90%" border="0" cellspacing="0" cellpadding="0" class="txt_white">
                                <% if rs.EOF = true then %>
										<br>
										<tr><td height="20">등록된 공지사항이 없습니다.</td></tr>
                                <% else %>
										<tr>
										<td height="10">&nbsp;</td>
										</tr>
									<% do while not rs.EOF %>
									<% i = 1 %>
	
										<tr>
										<td height="20">[공지]<a href="javascript:gonggi('<%=rs("no")%>')"><font color="#FFFF99"><% =rs("title") %></font></a></td>
										</tr>
									 <% rs.MoveNext 
									  response.flush %>
									<% loop %>
									
                                <% end if %>
                                
                                <!-- 모어 쓰실때 쓰시길..--><%' if i = 1 THEN %><!--a href="javascript:announclist()"--><%' end if%><!--more..-->
                              </table>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>			
          </td>
        </tr>
      </table>
<!--main img end-->				
				
				
				
				</td>
              </tr>
              <tr>
                <td>
				
	  <!--main-->		
	  <table width="580" border="0" cellspacing="12" cellpadding="0">
        <tr> 
          <td> 
<!--**************Best Collection 윗 상품 start***************************-->
			<table width="556" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="111" bgcolor="cccccc" valign="bottom" align="left"><img src="images_home/main_title01.gif" width="103" height="61"> 
                </td>
                
                <% do while not product_rs.EOF %>
                
                <td align="center" valign="top"> 
                  <table border="0" cellspacing="0" cellpadding="0" width="111">
                    <tr> 
                      <td align="center">
						
							<a href="catalogno.asp?catalogno=<% =product_rs("catalogno") %>">
						
						<img src="<% =product_rs("path") %>" border="0" width="90" height="90"></a>
						</td>
                    </tr>
                    <tr> 
                      <td align="center"valign="middle">
						
						  <a href="catalogno.asp?catalogno=<% =product_rs("catalogno") %>">
					
						
						<% =product_rs("catalognm") %></a>
					  </td>
                    </tr>
                    <tr>
						<td align="center" class="txt-66">
							<% if product_rs("maker") <> "" or isnull(product_rs("maker")) then  '제조사가 없는경우 공장그림 안나옴~%>
								<img src="images_home/icon_factory.gif" width="19" height="15" alt="제조사">&nbsp;<% =product_rs("maker") %>
							<% end if%>	
						</td>
                    </tr>
				  </table>
				</td>
				<% 
				   product_rs.movenext 
				   loop 
				%>

              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td>
<!--------------Best Collection 아래 상품 start-------------------------------------------->
			<table width="556" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<%' for i=1 to 5	
					do while not iventgood_rs_insert.EOF  %>
          			<td align="center" valign="top" width="111">
			            <table border="0" cellspacing="0" cellpadding="0" width="111">
			              <tr> 
			                <td align="center">
								
									<a href="catalogno.asp?catalogno=<% =iventgood_rs_insert("catalogno") %>">
							  	<img src="<% =iventgood_rs_insert("path") %>" width="90" height="90" border="0"></a></td>
			              </tr>
			              <tr>
			                <td align="center" class="txt-33">
								
									<a href="catalogno.asp?catalogno=<% =iventgood_rs_insert("catalogno") %>">
							
								<% =iventgood_rs_insert("catalognm") %></a></td>
			              </tr>
			              <tr> 
			                <td align="center" class="txt-66">
								<% if iventgood_rs_insert("maker") <> "" or isnull(iventgood_rs_insert("maker")) then  '제조사가 없는경우 공장그림 안나옴~%>
									<img src="images_home/icon_factory.gif" width="19" height="15" alt="제조사">&nbsp;<% =iventgood_rs_insert("maker") %>
								<% end if %>
							</td>
			              </tr>
			            </table>
			          </td>
					  <td>
							<%  'else
								 'end if
								 iventgood_rs_insert.MoveNext
								 loop
								 
								 'next %>
							<%' next %>
					  </td>
				</tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td background="images_home/dot.gif" height="1"></td>
        </tr>
        <tr> 
          <td> 
        
		
		
		
		
<!--**************new & cool 상품 start***************************-->
		<table width="556" border="0" cellspacing="0" cellpadding="0">
          <tr>
			<td width="112" bgcolor="cccccc" valign="bottom" align="left"><img src="images_home/main_title02.gif" width="77" height="69"> 
            </td>
            <td align="center" valign="top">
			
			
			<table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <% 
			    do while not iventgood_rs_new.EOF 
			    %>
				<td width="111" align="center" valign="top">
                  
				  
				  <!-- -------상품돌리는곳------------- -->
				  <table border="0" cellspacing="0" cellpadding="0" width="111">
                    <tr>
					  <td align="center">
                    	
							<a href="catalogno.asp?catalogno=<% =iventgood_rs_new("catalogno") %>">
					
                      		<img src="<% =iventgood_rs_new("path") %>" width="90" height="90" border="0"></a>
					  </td>
                    </tr>
                    <tr> 
                      <td align="center" class="txt-33">
                  
						<a href="catalogno.asp?catalogno=<% =iventgood_rs_new("catalogno") %>">
					
					<% =iventgood_rs_new("catalognm") %></a></td>
                    </tr>
                    <tr>
						<td align="center" class="txt-66">
							<% if iventgood_rs_new("maker") <> "" or isnull(iventgood_rs_new("maker")) then  '제조사가 없는경우 공장그림 안나옴~%>
								<img src="images_home/icon_factory.gif" width="19" height="15" alt="제조사">&nbsp;<% =iventgood_rs_new("maker") %>
							<% end if %>	
						</td>
                    </tr>
                  </table>
				  <!-- ----------------------------------- -->
				  
				  
                </td>
                 <%iventgood_rs_new.MoveNext %>
				 <%' next %> 
				 <% loop %>                
              </tr>
            </table>
			<!--new & cool-->
		  </td>
	    </tr>
	    </table>
<!--------------------------new & cool 상품 end--------------------------------->
				
				
				
		  </td>
        </tr>
		<tr>
		  <td valign="bottom">
<!--------------------------new & cool 상품 하단 Start--------------------------------->
			<table border="0" cellpadding="0" cellspacing="0" class="txt-33">
			<tr>
                <% 
			    do while not iventgood_rs_new_add.EOF 
			    %>
				<td width="112" align="center" valign="top">
                  
				  
				  <!-- -------상품돌리는곳_하단------------- -->
				  <table border="0" cellspacing="0" cellpadding="0" width="111">
                    <tr>
					  <td align="center">
                    	 
							<a href="catalogno.asp?catalogno=<% =iventgood_rs_new_add("catalogno") %>">
						
                      		<img src="<% =iventgood_rs_new_add("path") %>" width="90" height="90" border="0"></a>
					  </td>
                    </tr>
                    <tr> 
                      <td align="center" class="txt-33">
                  
						<a href="catalogno.asp?catalogno=<% =iventgood_rs_new_add("catalogno") %>">
					
					<% =iventgood_rs_new_add("catalognm") %></a></td>
                    </tr>
                    <tr>
						<td align="center" class="txt-66">
							<% if iventgood_rs_new_add("maker") <> "" or isnull(iventgood_rs_new_add("maker")) then  '제조사가 없는경우 공장그림 안나옴~%>
								<img src="images_home/icon_factory.gif" width="19" height="15" alt="제조사">&nbsp;<% =iventgood_rs_new_add("maker") %>
							<% end if %>	
						</td>
                    </tr>
                  </table>
				  <!-- ----------------------------------- -->
				  
				  
                </td>
                 <%iventgood_rs_new_add.MoveNext %>
				 <%' next 
				  response.flush
				  %> 
				 <% loop %>                
              </tr>
            </table>
			<!--new & cool-->
		  </td>
	    </tr>
		</table>
<!--------------------------new & cool 상품 end--------------------------------->
				
				
				
		  </td>
        </tr>
		<tr>
		  <td valign="bottom">		

     
		  </td>
		</tr>
      </table>
    </td>
  </tr>
</table>
    
	
	
	</td>
  </tr>
  <table width="750" border="0" cellspacing="0" cellpadding="0" height="100">
  <tr>
    <td><!--#include file="include/copyright.asp"--></td>
  </tr>
  </table>
</td>
<!--<td>
<div id="scrollmenu" style="left:880px; top:1px; z-index:10;">
	  <script type="text/javascript">
google_ad_client = "pub-7139684911602044";
google_ad_width = 120;
google_ad_height = 240;
google_ad_format = "120x240_as";
google_ad_type = "text";
google_ad_channel ="";
</script>
<script type="text/javascript"
  src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
</script>
</DIV>
</td>-->

</tr>
</table>
  </div>
</div>
</body>
</html>