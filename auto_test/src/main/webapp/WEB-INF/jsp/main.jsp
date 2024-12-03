<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
    <%@ page import="java.util.Date , java.text.SimpleDateFormat" %>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>개방해 OpenAPI 테스트</title>
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script>

$(function(){
    $("form").on("submit", function(event){
        event.preventDefault(); 
        
        
        var formData = new FormData(this); 
        
     	// 파일이 비어 있는지 확인
        var fileInput = $("#excelFile")[0]; // 파일 입력 필드
        if (fileInput.files.length === 0) {
            alert("엑셀 파일을 선택해주세요.");
            return;  // 파일이 없으면 업로드를 중지
        }
        
        
        $("#testBtn").attr("disabled",true);

        fetch('/uploadExcel.do', {
            method: 'POST',
            body: formData,  // formData에 파일 포함
        })
            .then(response => {
                if (response.ok) {
                    const contentType = response.headers.get("Content-Type");
                    if (contentType.includes("application/json")) {
                        return response.json().then(json => {
                            throw new Error(json.error);  // 서버에서 보낸 JSON 에러 메시지
                        });
                    } else if (contentType.includes("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
                        return response.blob();  // 엑셀 파일 다운로드 처리
                    }
                } else {
                    return response.text().then(text => {
                        throw new Error(text);  // 실패한 경우 응답 텍스트 반환
                    });
                }
            })
            .then(blob => {
                // Blob 형식으로 엑셀 파일 다운로드 처리
                if (blob.size > 0) {             
                    const link = document.createElement("a");
                    link.href = URL.createObjectURL(blob);
                    
                    const now = new Date();
                    const year = now.getFullYear().toString().slice(2); // 연도 마지막 두 자리
                    const month = String(now.getMonth() + 1).padStart(2, '0'); // 월 (0부터 시작하므로 +1)
                    const date = String(now.getDate()).padStart(2, '0'); // 일
                    const hours = String(now.getHours()).padStart(2, '0'); // 시
                    const minutes = String(now.getMinutes()).padStart(2, '0'); // 분
                    const seconds = String(now.getSeconds()).padStart(2, '0'); // 초
                    
                    const timestamp = "_"+year+month+date+hours+minutes+seconds;
                    
                 	link.download = "테스트 결과"+timestamp+".xlsx";  // 다운로드 파일명
                    link.click();
                } else {
                    alert("엑셀 파일 다운로드에 실패했습니다.");
                }
            })
            .catch(error => {
                alert(error.message || "파일 업로드에 실패했습니다.");
            })
            .finally(() => {
                $("#testBtn").attr("disabled", false);  // 버튼 활성화
            });
    });
});
</script>
</head>
<body>
<h3>개방해 OpenAPI 테스트</h3>
    <!-- 엑셀 파일 업로드 폼 -->
    <form name="popForm" method="post" enctype="multipart/form-data">
         <p><label for="excelFile">파일찾기</label></p>

         <p><input name="excelFile" id="excelFile" type="file" size="30" accept=".xls,.xlsx"></p>

        <!-- 업로드 버튼 -->
        <div class="btn-area">
            <button type="submit" id="testBtn" class="btn btn-yellow btn-ok">테스트</button>
           <!--  <button type="button" onclick="self.close();" class="btn btn-yellow btn-cancel">창닫기</button> -->
        </div>
    </form>
</body>
</html>