package egovframework.kss.main.service;

import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.commons.imaging.ImageReadException;
import org.apache.commons.imaging.Imaging;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import egovframework.kss.main.model.LayerData;

@Service
public class ExcelService {

	@SuppressWarnings("resource")
	public List<LayerData> parseExcelFile(MultipartFile file) throws Exception {
		List<LayerData> dataList = new ArrayList<>();

		if (file.isEmpty()) {
			throw new IllegalArgumentException("파일이 비어 있습니다.");
		}

		try (InputStream inputStream = file.getInputStream()) {
			Workbook workbook;

			// 파일 확장자에 따라 XSSFWorkbook (xlsx) 또는 HSSFWorkbook (xls) 사용
			if (file.getOriginalFilename().endsWith(".xlsx")) {
				try {
					if (inputStream.available() == 0) {
						throw new IllegalArgumentException("파일이 비어 있습니다.");
					}
					workbook = new XSSFWorkbook(inputStream);  // .xlsx 파일 처리
				} catch (Exception e) {
					throw new IllegalArgumentException("엑셀 파일(.xlsx)을 파싱하는 데 실패했습니다. 파일이 손상되었을 수 있습니다.");
				}
			} else if (file.getOriginalFilename().endsWith(".xls")) {
				try {
					workbook = new HSSFWorkbook(inputStream);  // .xls 파일 처리
				} catch (Exception e) {
					throw new IllegalArgumentException("엑셀 파일(.xls)을 파싱하는 데 실패했습니다. 파일이 손상되었을 수 있습니다.");
				}
			} else {
				throw new IllegalArgumentException("유효하지 않은 파일 형식입니다. 엑셀 파일을 업로드해 주세요.");
			}

			Sheet sheet = workbook.getSheetAt(0);  // 첫 번째 시트 선택

			Row headerRow = sheet.getRow(1); // 헤더 뽑아오기
			String layerNameHeader = getCellValue(headerRow.getCell(3)); // D열
			String layerEnglishHeader = getCellValue(headerRow.getCell(4)); // E열
			String url1Header = getCellValue(headerRow.getCell(12)); // M열
			String url2Header = getCellValue(headerRow.getCell(13)); // N열
			String url3Header = getCellValue(headerRow.getCell(14)); // O열
			String XMLHeader = getCellValue(headerRow.getCell(15)); // P열
			String JSONHeader = getCellValue(headerRow.getCell(16)); // Q열

			if (!layerNameHeader.equals("명칭") || !layerEnglishHeader.equals("레이어명") || !url1Header.equals("WMS 예시") || !url2Header.equals("WMS 이미지") || !url3Header.equals("WFS") || !XMLHeader.equals("XML") || !JSONHeader.equals("JSON")) { // 액셀 파일 형식 검사
				throw new Exception("액셀 파일의 양식이 다릅니다"); // 유효성 검사. 액셀 헤더가 양식과 다른 경우 예외 발생
			}

			// 데이터는 3번째 행 (인덱스 2)부터 시작
			for (int i = 2; i < sheet.getLastRowNum() && i < 52; i++) { // '&& i < ???'은 테스트시 api 전체가 아닌 일부만을 호출하기 위한 것으로, 불필요 할 시 삭제
				boolean comma = false; // 쉼표 찍을 건지
				Row row = sheet.getRow(i);
				if (row == null)
					continue;

				String layerName = getCellValue(row.getCell(3));  // D열
				String layerEnglishName = getCellValue(row.getCell(4));  // E열
				String url1 = getCellValue(row.getCell(12));  // K열
				String url2 = getCellValue(row.getCell(13));  // L열
				String url3 = getCellValue(row.getCell(14));  // M열
				String XMLUrl = getCellValue(row.getCell(15));  // L열
				String JSONUrl = getCellValue(row.getCell(16));  // M열
				String note = "";

				if (layerName.trim().isEmpty() || layerEnglishName.trim().isEmpty()) { //레이어의 명칭이나 레이어명이 없는 항목은 테스트하지 않고 넘어감. url 틀릴 경우 대비
					if (layerName.trim().isEmpty()) {
						note += "레이어 명칭 없음";
						comma = true;
					}

					if (layerEnglishName.trim().isEmpty()) {
						if (comma) {
							note += ", ";
						}
						note += "레이어명 없음";
					}
					dataList.add(new LayerData(i - 1, layerName, layerEnglishName, "", "", "", "", "", note)); // 검사 결과는 빈칸 적용
					continue;
				}

				// URL1 테스트 및 결과 저장
				url1 = !url1.trim().isEmpty() ? (callWMS(url1, layerEnglishName) ? "O" : "X") : "입력값 없음";

				// URL2 테스트 및 결과 저장
				url2 = !url2.trim().isEmpty() ? (callWMSImage(url2) ? "O" : "X") : "입력값 없음";

				// URL3 테스트 및 결과 저장
				url3 = !url3.trim().isEmpty() ? (callWFS(url3, layerEnglishName) ? "O" : "X") : "입력값 없음";

				XMLUrl = !XMLUrl.trim().isEmpty() ? (callXML(XMLUrl) ? "O" : "X") : "입력값 없음";

				JSONUrl = !JSONUrl.trim().isEmpty() ? (callJSON(JSONUrl) ? "O" : "X") : "입력값 없음";

				System.out.println(layerName + "= WMS : " + url1 + ", WMS 이미지 : " + url2 + ", WFS : " + url3 + ", XMLUrl : " + XMLUrl + ", JSONUrl : " + JSONUrl);

				dataList.add(new LayerData(i - 1, layerName, layerEnglishName, url1, url2, url3, XMLUrl, JSONUrl, note));
			}
		} catch (Exception e) {
			System.out.println("에러: " + e.getMessage());
			throw new IllegalArgumentException("엑셀 파일을 읽는 중 오류가 발생했습니다: " + e.getMessage());
		}

		return dataList;
	}

	private String getCellValue(Cell cell) {
		if (cell == null)
			return "";
		switch (cell.getCellType()) {
			case STRING:
				return cell.getStringCellValue();
			case NUMERIC:
				return String.valueOf(cell.getNumericCellValue());
			case BOOLEAN:
				return String.valueOf(cell.getBooleanCellValue());
			case FORMULA:
				return cell.getCellFormula();
			default:
				return "";
		}
	}

	//WMS 호출 검사 코드
	public boolean callWMS(String apiUrl, String layerCode) {
		HttpURLConnection connection = null;
		try {

			URL url = new URL(apiUrl);

			// HttpURLConnection 열기
			connection = (HttpURLConnection) url.openConnection();
			connection.setRequestMethod("GET");
			connection.setConnectTimeout(20000);
			connection.setReadTimeout(20000);

			// 요청 실행 및 응답 코드 확인
			int responseCode = connection.getResponseCode();
			if (responseCode != HttpURLConnection.HTTP_OK) {
				return false;
			}
			BufferedReader in = new BufferedReader(new InputStreamReader(connection.getInputStream()));
			StringBuilder response = new StringBuilder();
			String inputLine;
			while ((inputLine = in.readLine()) != null) {
				response.append(inputLine);
			}
			in.close();

			// 응답 본문 리턴
			String WMS_response = response.toString();

			String regex = "^var OtmsWmsLayer = new ol\\.source\\.TileWMS\\(\\{url:'http://www\\.khoa\\.go\\.kr/oceanmap/wmsdata\\.do', serverType:'mapserver', transition: 0, params:\\{ServiceKey:'[^']+',LAYERS:'" + Pattern.quote(layerCode) + "'\\}\\}\\);$";
			Pattern pattern = Pattern.compile(regex);
			Matcher matcher = pattern.matcher(WMS_response);

			// 매칭된 결과 반환
			return matcher.matches();

		} catch (Exception e) {
			System.err.println("API 호출 실패: " + apiUrl + " - " + e.getMessage());
			return false;
		} finally {
			if (connection != null) {
				connection.disconnect();  // 연결 해제
			}
		}
	}

	//WMS 이미지 호출 검사 코드
	public boolean callWMSImage(String apiUrl) {

		HttpURLConnection connection = null;
		try {

			URL url = new URL(apiUrl);

			// HttpURLConnection 열기
			connection = (HttpURLConnection) url.openConnection();
			connection.setRequestMethod("GET");
			connection.setConnectTimeout(20000);
			connection.setReadTimeout(20000);

			// 요청 실행 및 응답 코드 확인
			int responseCode = connection.getResponseCode();
			if (responseCode != HttpURLConnection.HTTP_OK) {
				return false;
			}

			// Content-Type 헤더 확인
			String contentType = connection.getContentType();
			if (contentType != null && contentType.startsWith("image/png")) {
				InputStream inputStream = connection.getInputStream();
				boolean isMonochrome = isImageMonochrome(inputStream); // 이미지가 단색인지 여부 판단

				if (isMonochrome) {
					System.out.println("이미지에서 레이어를 확인 할 수 없습니다.");
					return false; // 단색이라면 레이어가 찍혀있는 이미지가 아니므로 false 반환
				} else {
					System.out.println("이미지에 레이어가 존재합니다.");
					return true; // 레이어가 찍혀있는 이미지라면 true 반환
				}
			} else {
				return false; // png 파일이 아님
			}

		} catch (Exception e) {
			System.err.println("API 호출 실패: " + apiUrl + " - " + e.getMessage());
			return false;
		} finally {
			if (connection != null) {
				connection.disconnect();  // 연결 해제
			}
		}
	}

	// XML 파싱 방식 사용
	public boolean callWFS(String apiUrl, String layerCode) {
		try {
			// XML 파서 준비
			XMLReader xmlReader = XMLReaderFactory.createXMLReader();

			// 기본 핸들러 설정
			xmlReader.setContentHandler(new DefaultHandler() {
				@Override
				public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
					// "mapprime:" + layerCode.toLowerCase() 태그를 찾으면 파싱 종료
					if (qName.equalsIgnoreCase("mapprime:" + layerCode.toLowerCase())) {
						// 태그 발견 시 예외를 던져 파싱을 중단시킴
						throw new SAXException("레이어 발견 성공, 파싱 중지.");
					}
				}

				@Override
				public void endElement(String uri, String localName, String qName) throws SAXException {
					//  종료 처리 (필요시)
				}

				@Override
				public void endDocument() throws SAXException {
					// 문서가 끝났을 때 처리 (필요시)
				}
			});

			// URL에서 XML 데이터 읽기
			InputStream inputStream = new URL(apiUrl).openStream();
			xmlReader.parse(new InputSource(inputStream));

			// SAXException이 발생하지 않으면 레이어가 존재하지 않음
			return false;
		} catch (SAXException e) {
			return e.getMessage().contains("레이어 발견 성공");
		} catch (Exception e) {
			// 기타 예외 발생 시 처리
			System.err.println("API 호출 실패: " + apiUrl + " - " + e.getMessage());
			return false;
		}
	}

	// XML 호출.  XML 파싱 사용
	public boolean callXML(String apiUrl) {
		HttpURLConnection connection = null;
		try {

			DocumentBuilderFactory dbFactoty = DocumentBuilderFactory.newInstance();
			DocumentBuilder dBuilder = dbFactoty.newDocumentBuilder();
			Document doc = dBuilder.parse(apiUrl);

			doc.getDocumentElement().normalize();

			NodeList nList = doc.getElementsByTagName("item");
			System.out.println("파싱할 리스트 수 : " + nList.getLength()); // item 태그 개수 가져오기

			if (nList.getLength() > 0) {
				return true;
			} else {
				return false;
			}

		} catch (Exception e) {
			System.err.println("API 호출 실패: " + apiUrl + "-" + e.getMessage());
			return false;
		} finally {
			if (connection != null) {
				connection.disconnect();
			}
		}

	}

	public boolean callJSON(String jsonUrl) {
		HttpURLConnection connection = null;
		try {
			// URL 연결 설정
			URL url = new URL(jsonUrl);
			connection = (HttpURLConnection) url.openConnection();
			connection.setRequestMethod("GET");
			connection.setConnectTimeout(20000);
			connection.setReadTimeout(20000);

			// 응답 코드 확인
			int responseCode = connection.getResponseCode();
			if (responseCode == HttpURLConnection.HTTP_OK) {
				String contentType = connection.getHeaderField("Content-Type");
				if (contentType != null && contentType.contains("application/json")) {
					// JSON 데이터 읽기
					BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream()));
					StringBuilder response = new StringBuilder();
					String line;
					while ((line = reader.readLine()) != null) {
						response.append(line);
					}
					reader.close();

					// JSON 데이터 내부 값 확인
					return isJsonDataNonEmpty(response.toString());
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (connection != null) {
				connection.disconnect();
			}
		}
		return false;
	}

	private boolean isJsonDataNonEmpty(String jsonData) {
		try {
			// JSON 파싱
			ObjectMapper objectMapper = new ObjectMapper();
			JsonNode rootNode = objectMapper.readTree(jsonData);

			// 객체 또는 배열이 비어있는지 확인
			if (rootNode.isObject()) {
				return rootNode.size() > 0; // 객체에 속성이 있는지 확인
			} else if (rootNode.isArray()) {
				return rootNode.size() > 0; // 배열에 요소가 있는지 확인
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return false;
	}

	// 단색 이미지인지 확인하는 메소드
	public boolean isImageMonochrome(InputStream inputStream) throws IOException, ImageReadException {
		BufferedImage image = Imaging.getBufferedImage(inputStream);

		// 첫 번째 픽셀 색상 추출
		int firstPixelColor = image.getRGB(0, 0);

		// 이미지의 모든 픽셀을 순차적으로 확인... 매우 비효율적인 방식. 변경해야 할듯 => 생각보다 느리지 않다? wms 이미지 자체가 그리 크지 않게 와서 빠르게 처리된다.
		for (int y = 0; y < image.getHeight(); y++) {
			for (int x = 0; x < image.getWidth(); x++) {
				int pixelColor = image.getRGB(x, y);

				// 첫 번째 픽셀과 다른 색상이 존재하면 단색 아님
				if (pixelColor != firstPixelColor) {
					return false; // 단색이 아님
				}
			}
		}

		return true; // 모든 픽셀이 동일한 색상 => 단색
	}

}
