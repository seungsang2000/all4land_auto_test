package egovframework.kss.main.controller;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import egovframework.kss.main.model.LayerData;
import egovframework.kss.main.service.ExcelService;

@Controller
public class ExcelController {

	@Autowired
	private ExcelService excelService;

	/*@PostMapping("/uploadExcel.do")
	public void uploadExcel(@RequestParam("excelFile") MultipartFile file) {
		try {
			if (file == null || file.isEmpty()) {
				System.out.println("파일이 전달되지 않았습니다.");
				return;
			}
			System.out.println("파일 이름: " + file.getOriginalFilename());
			// 파일 처리
			List<LayerData> dataList = excelService.parseExcelFile(file);
			System.out.println(dataList); // 데이터를 출력해서 확인
	
		} catch (Exception e) {
			System.out.println("파일 처리 중 오류 발생: " + e.getMessage());
			e.printStackTrace();
		}
	}*/

	@PostMapping("/uploadExcel.do")
	public void uploadExcel(@RequestParam("excelFile") MultipartFile file, HttpServletResponse response) {
		try {
			if (file == null || file.isEmpty()) {
				System.out.println("파일이 전달되지 않았습니다.");
				return;
			}
			System.out.println("파일 이름: " + file.getOriginalFilename());

			// 엑셀 파일 파싱
			List<LayerData> dataList = excelService.parseExcelFile(file);
			System.out.println("엑셀 파일에서 읽은 데이터: " + dataList);

			// 엑셀 워크북 생성 (xlsx 포맷)
			Workbook workbook = new XSSFWorkbook();  // XSSFWorkbook은 .xlsx 형식
			Sheet sheet = workbook.createSheet("Layer Data");

			// 헤더 생성
			Row headerRow = sheet.createRow(0);
			headerRow.createCell(0).setCellValue("Layer Name");
			headerRow.createCell(1).setCellValue("Layer English Name");
			headerRow.createCell(2).setCellValue("URL 1");
			headerRow.createCell(3).setCellValue("URL 2");
			headerRow.createCell(4).setCellValue("URL 3");

			// 데이터 추가
			for (int i = 0; i < dataList.size(); i++) {
				LayerData data = dataList.get(i);
				Row dataRow = sheet.createRow(i + 1);
				dataRow.createCell(0).setCellValue(data.getLayerName());
				dataRow.createCell(1).setCellValue(data.getLayerEnglishName());
				dataRow.createCell(2).setCellValue(data.getUrl1());
				dataRow.createCell(3).setCellValue(data.getUrl2());
				dataRow.createCell(4).setCellValue(data.getUrl3());
			}

			// 엑셀 파일을 서버의 로컬 디스크에 저장 (Optional)
			try (FileOutputStream fileOut = new FileOutputStream("C:/Temp/processedLayerData.xlsx")) {
				workbook.write(fileOut);
			}
			System.out.println("엑셀 파일이 서버에 저장되었습니다. 파일 경로: C:/Temp/processedLayerData.xlsx");

			// 파일 다운로드 응답 설정
			response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
			response.setHeader("Content-Disposition", "attachment; filename=processedLayerData.xlsx");

			// 엑셀 파일을 HTTP 응답으로 전송
			try (ServletOutputStream outputStream = response.getOutputStream()) {
				workbook.write(outputStream);
				outputStream.flush();  // 스트림을 플러시하여 데이터를 완전히 전송
			} catch (IOException e) {
				System.out.println("파일 스트림을 작성하는 중 오류 발생: " + e.getMessage());
				e.printStackTrace();
			} finally {
				workbook.close();  // 워크북을 닫아서 리소스를 해제
			}

		} catch (Exception e) {
			System.out.println("파일 처리 중 오류 발생: " + e.getMessage());
			e.printStackTrace();
		}
	}

}
