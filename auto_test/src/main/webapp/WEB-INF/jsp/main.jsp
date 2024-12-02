<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Insert title here</title>
</head>
<body>
	<form name="popForm" method="post" action="/uploadExcel.do" enctype="multipart/form-data">
				<table>
				<caption>엑셀 업로드</caption>
				<colgroup>
				<col><col>
				</colgroup>
				<tbody>
				<tr>
					<th><label for="code2">파일찾기</label></th>
					<td><input name="excelFile" id="excelFile" type="file" size="30"></td>
				</tr>
				</table>
				<!-- 버튼 영역 -->
				<div class="btn-area">
				    <button type="submit" class="btn btn-yellow btn-ok">업로드</button>
				    <button type="button" onclick="self.close();" class="btn btn-yellow btn-cancel">창닫기</button>
				</div>
			</form>
</body>
</html>