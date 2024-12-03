package egovframework.kss.main.controller;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
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
	@ResponseBody
	public void uploadExcel(@RequestParam("excelFile") MultipartFile file, HttpServletResponse response) {
		try {
			if (file == null || file.isEmpty()) {
				response.setStatus(HttpServletResponse.SC_BAD_REQUEST);  // 400 오류
				response.setContentType("application/json; charset=UTF-8");
				response.getWriter().write("파일이 전달되지 않았습니다.");
				return;
			}
			System.out.println("파일 이름: " + file.getOriginalFilename());

			// 엑셀 파일 파싱
			List<LayerData> dataList = excelService.parseExcelFile(file);
			System.out.println("엑셀 파일에서 읽은 데이터: " + dataList);

			// 엑셀 워크북 생성 (xlsx 포맷)
			Workbook workbook = new XSSFWorkbook();  // XSSFWorkbook은 .xlsx 형식
			Sheet sheet = workbook.createSheet("Layer Data");

			// 헤더 셀 스타일 생성
			CellStyle headerStyle = workbook.createCellStyle();

			// 배경색을 회색으로 설정
			headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			sheet.createRow(0);

			// 헤더 생성
			Row headerRow = sheet.createRow(1);
			Cell cell0 = headerRow.createCell(1);
			cell0.setCellValue("번호");
			cell0.setCellStyle(headerStyle); // setCellStyle은 메소드 체이닝이 안되므로 따로 분리

			Cell cell1 = headerRow.createCell(2);
			cell1.setCellValue("명칭");
			cell1.setCellStyle(headerStyle);

			Cell cell2 = headerRow.createCell(3);
			cell2.setCellValue("레이어명");
			cell2.setCellStyle(headerStyle);

			Cell cell3 = headerRow.createCell(4);
			cell3.setCellValue("WMS");
			cell3.setCellStyle(headerStyle);

			Cell cell4 = headerRow.createCell(5);
			cell4.setCellValue("WMS 이미지");
			cell4.setCellStyle(headerStyle);

			Cell cell5 = headerRow.createCell(6);
			cell5.setCellValue("WFS");
			cell5.setCellStyle(headerStyle);

			Cell cell6 = headerRow.createCell(7);
			cell6.setCellValue("비고");
			cell6.setCellStyle(headerStyle);

			// 데이터 추가
			for (int i = 0; i < dataList.size(); i++) {
				LayerData data = dataList.get(i);
				Row dataRow = sheet.createRow(i + 2);
				dataRow.createCell(1).setCellValue(data.getOrder());
				dataRow.createCell(2).setCellValue(data.getLayerName());
				dataRow.createCell(3).setCellValue(data.getLayerEnglishName());
				dataRow.createCell(4).setCellValue(data.getUrl1());
				dataRow.createCell(5).setCellValue(data.getUrl2());
				dataRow.createCell(6).setCellValue(data.getUrl3());
				dataRow.createCell(7).setCellValue(data.getNote());
			}

			// 각 열에 대해 자동 크기 조정
			for (int i = 0; i <= 7; i++) {
				sheet.autoSizeColumn(i);
			}

			// 엑셀 파일을 서버의 로컬 디스크에 저장 (선택)
			try (FileOutputStream fileOut = new FileOutputStream("C:/Temp/TestData.xlsx")) {
				workbook.write(fileOut);
			}
			System.out.println("엑셀 파일이 서버에 저장되었습니다. 파일 경로: C:/Temp/TestData.xlsx");

			// 파일 다운로드 응답 설정
			response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
			response.setHeader("Content-Disposition", "attachment; filename=TestData.xlsx");

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
			response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);  // 500 오류
			response.setContentType("application/json; charset=UTF-8");
			try {
				response.getWriter().write(e.getMessage());
			} catch (IOException e1) {
				e1.printStackTrace();
			}
		}
	}

}
