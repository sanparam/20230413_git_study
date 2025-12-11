<%@ Language=VBScript %>
<!--#include file="./sql_injection.asp"-->
<% Response.buffer=true 
	Response.Expires =0
	
	ls_catalogno = Request ("catalogno")
	page = request("page")
	code = request("code")
	gubun= request("gubun")


Set cn = Server.CreateObject("ADODB.Connection")
	cn.Open Application("db_ConnectionString")
	
Set pend_good = Server.CreateObject("ADODB.Recordset") 
	querystr = "select g.catalogno, g.goodcd,g.modelnm,g.pend from good g where g.catalogno = '" & ls_catalogno & "'"
    pend_good.ActiveConnection = cn
    pend_good.CursorType = adOpenStatic
    pend_good.Source = querystr
    pend_good.Open

Set catalog = Server.CreateObject("ADODB.Recordset") 
	querystr = "select c.catalogno, c.catalognm, c.keyfield, c.vat, c.maker, c.formtype, di.path as img_path  ,"
	querystr = querystr & " dt.title , dt.stext, c.moneynm "
	querystr = querystr & "from catalog c left join dicatalog di on c.catalogno = di.catalogno  "
	querystr = querystr & " left join dtcatalog dt on c.catalogno = dt.catalogno "
	querystr = querystr & " where  c.catalogno = '" & ls_catalogno & "'"
    catalog.ActiveConnection = cn
    catalog.CursorType = adOpenStatic
    catalog.Source = querystr
    catalog.Open
    
set df_catalog = server.CreateObject ("adodb.recordset")
	df_catalog.ActiveConnection = cn
	sql ="select title, path as file_path from dfcatalog where catalogno = '"& ls_catalogno &"' "
	df_catalog.Open sql

'set df_catalog_1 = server.CreateObject ("adodb.recordset")
	'df_catalog_1.ActiveConnection = cn
	'sql ="select path as file_path, Title from dfcatalog where catalogno = '"& ls_catalogno &"' "
	'df_catalog_1.Open sql

set supplier = server.CreateObject ("adodb.recordset")
	supplier.ActiveConnection = cn
	sql ="select * from catalog where supply_code = '"& session("code") &"' and catalogno ='" & ls_catalogno & "'"
	supplier.Open sql
%>

<html>
<head>
	<title>Macro21 Co. &nbsp;<%=catalog("catalognm")%></title>
	<link rel="stylesheet" href="pjdoc.css" type="text/css" title="macro21 css">
	<meta property="fb:admins" content="&#123;macro21MRO&#125;"/>
</head>
<script language="javascript">
function js_image_view(image_file) {
	window.open('/search/bigimage.asp?arg='+image_file, 'message', 'width=500, height=500, scrollbars=1, resizable=1');
}

function send(temp) {
	
	var frm=document.i_ck;
	var str;
	
	if ( frm.ref.length > 0 ) {
		i = 0;
		b_cycle = false;
		while ( i<= frm.ref.length-1 && b_cycle == false ) {
		if ( frm.ref[i].value == 'T' ) {
				b_cycle = true
			
				if (frm.tqt[i].value < frm.limitqty[i].value ) {
					alert ("최소주문수량보다 작습니다.1");
					frm.tqt[i].focus();
					
					return;
				}
				
							
			} else {
			}
			i = i + 1
		}
		if ( b_cycle ) {
			
		}
		else {
			alert ( "제품을 선택하세요 ..." );						
			return;			
		}
	
	} else {
		
		if ( frm.ck_1.checked ) {
		
			if (frm.tqt.value < frm.limitqty.value ) {
				alert("최소주문수량보다 작습니다.");
				return;
			}
		}
		else {
			alert ( "제품을 선택하세요 ..." );			
			return;
		}
	}
	frm.status.value = temp
	frm.action = 'commerce/sd_make.asp';
	
	
	frm.method='post'
	frm.submit();
}


function js_check(i){
	var frm=document.i_ck;
	
	if ( frm.ref.length > 0 ) {
		if ( frm.ref[i-1].value == 'T' ) {
			frm.ref[i-1].value = 'F';
		} else {
			frm.ref[i-1].value = 'T';		
		}
		return;	
	} else {
		if ( frm.ck_1.checked ) {
			frm.ref.value = 'T';	
		} else {
			frm.ref.value = 'F';		
		}
		return;					
	}        
}

function js_check_1(i){
	var frm=document.i_ck;
	if ( frm.ck_1.checked ) {
	
	alert("취급중단된 품목입니다.");
	history.back(+1);	
					
	}        
}

</script>

<body bgcolor="#ffffff" topmargin="20" leftmargin="0" marginheight="20" marginwidth="0" vlink="#8e8e8e" link="#8e8e8e" alink="#8e8e8e">

<table border=0 cellpadding=0 cellspacing=0 width=700>
<tr><td><font color="red"><b>&nbsp;&nbsp;&nbsp;▦ Catalog No. : M-<%=ls_catalogno%></b></font></td></tr>                 
  <%'-------------------------------------------------                 
	' 분류 보이기                 
	'-------------------------------------------------                 
	set class_rs = Server.CreateObject("ADODB.Recordset")                 
		class_rs.ActiveConnection = cn                 
		                 
		if code  = "" then                 
		sql = "select * from corp_class c join corp_catalog cc on c.code = cc.corp_class_code " & _                 
	     "where cc.catalogno = '" & ls_catalogno & "' "                 
	     else                 
	     sql = "select * from corp_class c  " & _                 
	           "where c.code = " & code & " "                 
	                      
		end if                 
		'Response.Write sql                 
		class_rs.Open sql                 
	                 
	if class_rs.EOF = false then                 
		sortname = class_rs("full_name")                 
		                 
		sortname = split(sortname,"&gt;") %>                 
		                 
		<tr>                 
			<td colspan=2 class=bodyblack align=right valign=bottom>                 
				<img src="../images03/point.gif" width="9" height="9" alt="" border="0">&nbsp;                 
				<a href="/index_body.asp">Home</a> &gt;                  
					<% for ii = 0 to ubound(sortname) -1 %>                 
						<% if ii = ubound(sortname) - 2 then %>                 
							<% if gubun="image" then %>                 
								<a href="search/category_img.asp?code=<% =class_rs("parent_code") %>"><% =sortname(ii) %></a> &gt;                 
							<% else %>                 
								<a href="search/category.asp?code=<% =class_rs("parent_code") %>"><% =sortname(ii) %></a> &gt;                 
							<% end if %>                 
						<% elseif ii = ubound(sortname) - 1 then %>                 
							<% if gubun="image" then %>                 
								<a href="search/category_img.asp?code=<% =class_rs("code") %>&page=<%=page%>"><% =sortname(ii) %></a> &gt;                 
							<% else %>                 
								<a href="search/category.asp?code=<% =class_rs("code") %>&page=<%=page%>"><% =sortname(ii) %></a> &gt;                 
							<% end if %>                 
						<% else                  
					                  
							sql = "select * from corp_class where code = '" & class_rs("parent_code") & "'"                   
							'Response.Write  sql                  
							set sub_rs1 = cn.Execute (sql)                 
							                  
							if not sub_rs1.eof then %>                 
								<% if gubun="image" then %>                 
									<a href="search/category_img.asp?code=<% =sub_rs1("parent_code") %>&page=<%=page%>"><% =sortname(ii) %></a> &gt;                 
								<% else %>                 
									<a href="search/category.asp?code=<% =sub_rs1("parent_code") %>&page=<%=page%>"><% =sortname(ii) %></a> &gt;                 
								<% end if %>                 
						    <%end if%>                 
						<% end if %>                 
					<% next %>                 
			                 
			</td>                 
		</tr>                 
		                 
	<%                 
	                 
	end if                 
	'-------------------------------------------------                 
	' end : 분류 보이기                 
	'-------------------------------------------------                 
	%>                 
  <tr>                 
    <td colspan=2><img src="../images03/dot.gif" width="700" height="2" alt="" border="0"></td>                 
  </tr>                 
  <tr><td colspan=2 height=5>&nbsp;</td></tr>                 
  <tr>                 
    <td width=300 align=center valign=top>                 
	<!--- 제품 이미지 --->                 
	<% if catalog("img_path") <> "" then  
	
	capital_path = catalog("img_path")

	file_type = Mid(capital_path, 54,4)
	'Response.write file_type
	'Response.write "<br>"
		If file_type = "file" then
		dfile_no = Mid(capital_path, 59,4)
		'Response.write dfile_no
		'Response.write "<br>"
		int_dfile_no = CInt(dfile_no)
		'Response.write int_dfile_no
		'Response.write "<br>"
			If int_dfile_no < 0576 Then
				
			%>
				<!--<a href="javascript:js_image_view('<% =capital_path %>');"><img SRC="<% =capital_path%>" width="200" height="200" alt="" border="1"
				 onerror="this.src='/images_home/icon_factory.gif';"></a>//--> 
				<img SRC="<% =capital_path%>" width="200" height="200" alt="" border="1"
				 onerror="this.src='/images_home/icon_factory.gif';">      
			<br>                 
				
			<% Else
				capital_path = replace(capital_path, "image", "Image") %>
				<!--<a href="javascript:js_image_view('<%=capital_path%>');"><img src="<%=capital_path%>" width="200" height="200" alt="" border="1" onerror="this.src='/images_home/icon_factory.gif';"></a>//-->  
				<img SRC="<% =capital_path%>" width="200" height="200" alt="" border="1"
				 onerror="this.src='/images_home/icon_factory.gif';"> 
				<br> 
				&nbsp;                 
			<% end if %>
	 <% Else %>
		<!--<a href="javascript:js_image_view('<% =capital_path %>');"><img SRC="<% =capital_path%>" width="200" height="200" alt="" border="1"
		onerror="this.src='/images_home/icon_factory.gif';"></a> //-->
				<img SRC="<% =capital_path%>" width="200" height="200" alt="" border="1"
				 onerror="this.src='/images_home/icon_factory.gif';"> 
				<br>                 
		
	 <% End If %>

	<!-- <a href="javascript:js_image_view('<%=capital_path %>');"><img src="../images03/magnify.gif" width="95" height="24" alt="" border="0"></a> //-->              
			
 <% End If %>

	</td>                 
	<td valign=top align=center>                 
	  <table border=0 cellpadding=3 cellspacing=0 width=450>                 
	    <tr>                 
		  <td class=subtitlebold height="22"><!--- 제품명 ---><%=catalog("catalognm")%></td>                 
		</tr>                 
		<tr>
		  <td>
		    <!--- 제조사/배송방법/결제방법/첨부파일 테이블 시작 --->
		    <table border=0 cellpadding=0 cellspacing=0>
			  <tr>
			    <td><img src="images03/table_top1.gif" width="442" height="21" alt="" border="0"></td>
			  </tr>
			  <tr>
			    <td background="images03/back_white.gif" class=bodyheight>&nbsp;&nbsp;<font color="#808080">제조사 :</font> &nbsp;&nbsp;<%=catalog("maker")%></td>
			  </tr>
			  <tr>
			    <td><img src="images03/dot_gray.gif" width="442" height="1" alt="" border="0"></td>
			  </tr>
			  <tr>
			    <td background="images03/back_gray.gif" class=bodyheight>&nbsp;&nbsp;<font color="#808080">Key word :</font> &nbsp;&nbsp;<%=catalog("keyfield")%></td>
			  </tr>
			  <tr>
			    <td><img src="images03/dot_gray.gif" width="442" height="1" alt="" border="0"></td>
			  </tr>
			    <!--<tr>
			    <td background="images03/back_white.gif" class=bodyheight>&nbsp;&nbsp;<font color="#808080">결제방법 :</font> &nbsp;&nbsp;일괄결제</td>
			  </tr>-->
			  <tr>
			    <td><img src="images03/dot_gray.gif" width="442" height="1" alt="" border="0"></td>
			  </tr>
			  <tr>
			    <td background="images03/back_gray.gif" class=bodyheight>&nbsp;&nbsp;<font color="#808080">첨부파일 :</font> &nbsp;&nbsp;
					<% do while not df_catalog.EOF %>
						&nbsp;&nbsp;&nbsp;<li><a href="<%=df_catalog("file_path")%>"><img src="images03/disk.gif" width="13" height="13" alt="" border="0">&nbsp;<%=df_catalog("title")%></a>
					<% df_catalog.MoveNext()
					   loop
					%></td>
			  </tr>
			  <tr>
			    <td><img src="images03/bottom_gray.gif" width="442" height="2" alt="" border="0"></td>
			  </tr>
			</table>
			<!--- 제조사/배송방법/결제방법/첨부파일 테이블 끝 --->
			<br><br>
			
		  </td>
		</tr>
	 </table>
	</td>
  </tr>
  <tr>
    <td colspan=2><img src="images03/dot.gif" width="720" height="2" alt="" border="0"></td>
  </tr>   
                      
  <tr>                 
	<td colspan=2 class=bodyblack>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images03/spec.gif" width="141" height="24" alt="" border="0">                 
				&nbsp;&nbsp;&nbsp;&nbsp;아래 가격은                 
				<font color=blue> <% if catalog("vat") = "Y" then %>                 
					 VAT 포함                 
				<% else %>                 
					VAT 별도                 
				<% end if %>&nbsp; 참고용임</font>                 
    </td>                 
  </tr>                 
  <%                 
                   
  Set HeadInfo = Server.CreateObject("ADODB.Recordset")                  
	  HeadInfo.ActiveConnection = cn                 
      HeadInfo.CursorType = adOpenStatic                 
      sql ="select title, subcount, catalogno, no from cataloghead where headlevel='1' and catalogno ='"& ls_catalogno &"' "                 
      sql = sql & " order by no "                 
      HeadInfo.Open sql                 
                     
  %>                 
  <% if catalog("formtype")= "1" then %>                 
  <tr>                 
    <td colspan=2 align="center">                 
	  <!--- 제품사양 테이블 시작 --->                 
	  <table border=0 cellpadding=3 cellspacing=1 width=720 bgcolor="#aaaaaa">                 
	    <tr>                 
		  <td bgcolor="#dbdbdb" class=bodyblack align=center width='10%'>제품명</td>                 
		  <td bgcolor="#dbdbdb" class=bodyblack align=center width='5%'>규격</td>                 
		                  
		   <%	j = 0 ' sub column count location ...
				n = 0 ' ?
				redim Preserve arr(1)
				arr(0) = 0
				'arr(1) = 0
				do while not headinfo.eof
					j = j + 1
					if headinfo.fields(1) = "0" then    'subcount = "0" 
						redim Preserve arr(j + 1)
						arr(j) = 0
						Response.write "<td bgcolor='#dbdbdb' class=bodyblack align=center width='5%'>" & headinfo.fields(0) & "</td>"
						n = n + 1		
					else
						Response.write "<td bgcolor='#dbdbdb' class=bodyblack align=center width='5%'>" & headinfo.fields(0)
						'sh.Source = "exec proc_catalog_11 '" & headinfo.fields(2) & "', '" & headinfo.fields(3) & "'"
						sql ="select title, subcount, catalog from cataloghead where catalogno = '"& ls_catalogno &"' and m_headno = '" & headinfo.fields(3) & "' and not ( HEADLEVEL = '1' ) order by  m_headno,no"
						sh.Open			
						Response.write "<table width='100%' border='0' cellspacing='1' cellpadding='0'>"
						Response.write "<tr align='left' class='bodyblack'>"
						k = 0 ' sub count
						do while not sh.eof
							k = k + 1
							Response.write "<td bgcolor='#ffffdd' width='5%' class='bodyblack'>" & sh.Fields(0) & "</td>"
							sh.movenext
							n = n + 1
						loop
						redim Preserve arr(j + 1)
						arr(j) = k
						Response.write "</tr>"
						Response.write "</table>"
						sh.close
						Response.write "</td>"					
					end if
					headinfo.movenext
				loop
			%>
		     
		                   
		  <td bgcolor="#dbdbdb" class=bodyblack align=center width='5%'>가격</td>                 
		  <td bgcolor="#dbdbdb" class=bodyblack align=center width='7%'>최소주문수량</td>                 
		  <% If session("code") = "" Then %>                 
		  <td bgcolor="#dbdbdb" class=bodyblack align=center width='3%'>선택</td>                 
		  <% End If %>                 
		</tr>                 
		<form NAME="i_ck">                 
			<input type="hidden" name="status">                 
			<input type="hidden" name="catalogno" value="<%=ls_catalogno%>">                 
			<input type="hidden" name="code" value="<%=code %>">                 
			<input type="hidden" name="page" value="<%=page%>">                 
			                 
			<%                 
				Set SizeInfo = Server.CreateObject("ADODB.Recordset")                  
					querystr = "exec proc_catalognew_12 '" & ls_catalogno & "'"                 
					SizeInfo.ActiveConnection = cn                 
					SizeInfo.CursorType = adOpenStatic                 
					SizeInfo.Source = querystr                 
					SizeInfo.Open                 
					                 
				li_loop = 0                 
				no_order_flag = true		' 납기나 금액이 없다면 [발주]를 나타내지 않기 위해 사용된 flag                 
				                 
				                 
				                 
				do while not sizeinfo.eof                 
				Response.write "<tr>"                 
				i = 0                 
			    g = 0                 
			                     
			    Response.write "<td bgcolor='#ffffff' class=bodyblack align=center width='5%' nowrap>" & sizeinfo(1) & "</td>"  	  								                 
				Response.write "<td bgcolor='#ffffff' class=bodyblack align=center width='5%' nowrap>" & sizeinfo(2) & "</td>"  	  								                 
				                 
			                     
			    do while i < n                 
			                     
			    	if sizeinfo(i+3) = "" then                 
			    		temp_str = "·"                 
			    	else                 
			    		temp_str = sizeinfo(i+3)                  
			    	end if                 
			    	                 
			    	if arr(g) = 0 then                 
						Response.write "<td bgcolor='#ffffff' class=bodyblack align=center width='5%' nowrap>" & temp_str & "</td>"  	  								                 
						i = i + 1 		  					                 
			  		else                 
						Response.write "<td bgcolor='#ffffff' class=bodyblack align=center >"			                 
						Response.write "<table width='100%' border='0' cellspacing='1' cellpadding='0'>"                 
						Response.write "<tr align='center' class='bodyblack'>"			                 
						for m = 1 to arr(g)                 
						Response.write "<td width='" & 100/arr(g) & "%' class='bodyblack' bgcolor='#F6F6F6' nowrap>" & temp_str & "</td>"  	  					                 
						i = i + 1			                 
						next  		                 
						Response.write "</tr>"                 
						Response.write "</table>"  		                 
						Response.write "</td>"			  		                 
			  		end if                 
					g = g + 1                 
			  	loop                 
				                 
				li_loop = li_loop + 1                 
				ls_loop = cstr( li_loop )	  	                 
                 
				                 
				if sizeinfo.Fields("b").value = "0" or sizeinfo.fields("b") = ""  or isnull(sizeinfo.fields("b")) then                 
				   Response.write "<td valign='middle' align=center width='5%' class='bodyblack' bgcolor='#ffffff' align='center'> " & "견적가" & "</td>"			                 
				   no_order_flag = false		' 금액이 없다면 [발주]를 나타내지 않기 위해 사용된 flag                 
				else                 
					Response.write "<td valign='middle' width='5%' class='bodyblack' bgcolor='#ffffff' align='center'>" &  formatnumber(sizeinfo.Fields("b").value,0) & catalog("moneynm") & "</td>"	                 
				end if                 
			                 
				Response.write "<td valign='middle' align='center' width='7%' class='bodyblack' bgcolor='#ffffff'>"		                 
				Response.Write "(" & sizeinfo.Fields("a").Value &  ")" %>
				<INPUT type="text" name="tqt" value=1 maxlength="7" size="7" style="text-align:right;"></input>                 
				
				<% Response.write "</td>"                 
				                 
				If session("code") = "" then		                 
				Response.write "<td valign='middle' width='3%' class='bodyblack' bgcolor='#ffffff' align='center'>"                 
				                 
				  if pend_good("pend") = "Y" then                 
				Response.Write "<INPUT type='checkbox'"                  
				response.write "disabled=True"                 
				response.write " name='ck_" & ls_loop & "' onclick='js_check_1(" & ls_loop & ")'><br><font color='red'>취급중단</font></td>"                 
				Else                 
				Response.Write "<INPUT type='checkbox' name='ck_" & ls_loop & "' onclick='js_check(" & ls_loop & ")'></td>"                 
				end if                  
				                 
				Response.Write "<INPUT type='hidden' name='goodcd_" & ls_loop & "' value='" & sizeinfo.Fields(0).value & "'>"                 
				 End If                 
				                 
				Response.write "<INPUT type='hidden' name='ref' value='F'>"	                 
				Response.Write "<input type='hidden' name='limitqty' value='" & sizeinfo.Fields("a").Value & "'>"                 
				Response.write "<INPUT type='hidden' name='goodnm_" & ls_loop & "' value='" + sizeinfo.Fields(1).value + "'>"                 
			  	Response.write "</tr>"                 
				sizeinfo.movenext                 
                 
			loop                   
			Response.Write "<INPUT type='hidden' name='rowcount' value='" & ls_loop & "'>"                 
			%>
			</form>
	  </table>                 
	  <!--- 제품사양 테이블 끝 --->                 
	</td>                 
  </tr>                   
  <% end if %>                 
</table>                 
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
   <tr align="left" valign="middle"> 
	<td style="padding-left:250px" height="17">
	[<a href="javascript:history.back(-1)">뒤로</a>]	                 
	                 
	<% if not supplier.EOF then%>
	      [<a href="/history/default.asp?main=m&catalogno=<%=ls_catalogno%>">제품 수정</a>]
    <%end if%>
    
    <% if no_order_flag = true  then %>
          [<a href="javascript:send('request');">주  문</a>]
    <% end if %>
    
		  [<a href="javascript:send('quote');">견적 요청</a>]                 
</td>                 
</tr>                
<tr>
<!--- 상세정보 테이블 시작 --->
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<table border=0 cellpadding=0 cellspacing=20>
			  <tr>
			    <td><img src="images03/detail_top.gif" width="442" height="21" alt="" border="0"></td>
			  </tr>
			  <tr>
			    <!--<td background="images03/back_white.gif">//-->
				<td>
				  <table border=0 cellpadding=10 cellspacing=0 width=100% class=bodyheight>
				    <tr>
					  <td>
					  <b><%=catalog("title")%></b>
					  <br>
					  <%=trim(catalog("stext"))%>
					  </td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr>
			    <td><img src="images03/bottom_gray.gif" width="442" height="2" alt="" border="0"></td>
			  </tr>
			</table>
			
</tr>
<!--- 상세정보 테이블 끝 --->
<tr align="center"><td> <font color="blue"><b>&nbsp;&nbsp;&nbsp;Twitter Q&A ;</b></font>&nbsp;<a href="http://twitter.com/share" class="twitter-share-button" 
 	data-count="none" data-via="windfrommount" data-related="windfrommount:option"></a>
 	<script async src="https://platform.twitter.com/widgets.js" charset="utf-8"></script>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
 	</td>
	
	<script language="javascript">
 					function kakaoOpen()

					{
       				window.open("https://open.kakao.com/o/s7Lpphxc", "newWin");
					} 
			</script> 
	<!--<td>
	<a href="https://open.kakao.com/o/s7Lpphxc"><font color="blue"><b>&nbsp;&nbsp;&nbsp;KakaoTalk</b></font></a></td>-->
	<td>
	<font color="red" onClick="kakaoOpen();" style="cursor:pointer;"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;KakaoTalk open CHAT !!!</b></td>


<br><br>
<td> <font color="blue"><b>&nbsp;&nbsp;&nbsp;&nbsp;Facebook Comment ;</b></font>
<div class="fb-messengermessageus" 
			  messenger_app_id="1775262505983580" 
			  page_id="577906588888452"
			  color="blue"
			  size="standard">
			</div>
</td>				
</tr>
</table>
<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="20">
		 <tr>
		 <td>
		    <div id="fb-root"></div>
				<script>(function(d, s, id) {
				var js, fjs = d.getElementsByTagName(s)[0];
				if (d.getElementById(id))
				return;
				js = d.createElement(s);
				js.id = id;
				js.src = "//connect.facebook.net/ko_KR/all.js#xfbml=1&appId=177592243667231";
				fjs.parentNode.insertBefore(js, fjs);
				}
				(document, 'script', 'facebook-jssdk')
				);
				</script>
			
            <div class="fb-comments" data-href="http://www.macro21.com/catalogno.asp?catalogno=<%=ls_catalogno%>" data-width="550" data-num-posts="2" order_by="reverse_time"></div>
			<!--<div class="fb-page" data-href="https://www.facebook.com/macro21MRO" data-tabs="timeline" data-width="" data-height="" data-small-header="false" data-adapt-container-width="true" data-hide-cover="false" data-show-facepile="true"><blockquote cite="https://www.facebook.com/macro21MRO" class="fb-xfbml-parse-ignore"><a href="https://www.facebook.com/macro21MRO">Facebook</a></blockquote></div>-->
			   
				
			
         </td>
		 
		 
		 </tr>
		</table>
</td>
</tr>
<table border=0 cellpadding=0 cellspacing=0 width=720>
  <tr>
    <td align="center"><hr size="1"></td>
  </tr>
  <tr>	
    <td align=center>
	<% If InStr(class_rs("full_name"),"temp_140627 글자수보다 크다")> 0 Then %>
	<!--#include file="include/macro21_1.htm"-->
	<% Else %>
	<!--#include file="include/macro21.htm"-->
	<% End If %>
	</td>
	<% if session("userlevel") = "7" Then %>

	       <script language="javascript">
 					function winOpen()

					{
       				window.open("mcatalogno.asp?catalogno=<%=ls_catalogno%>", "newWin");
					} 
			</script> 
    <td> <div class="fb-like" data-send="true" data-width="700" data-show-faces="true"></div>
			   <br>
	</td>
	<td><font onClick="winOpen();">Transfer</font></td>
	<% end if %>
  </tr>
</table>
<td align=center><!--#include file="ad_plan_r1.asp"--></td>

</body>
</html>