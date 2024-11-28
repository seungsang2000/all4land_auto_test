package egovframework.kss.main;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
public class mainController {

	@RequestMapping("/main.do")
	public String mainPage() {
		System.out.println("테스트테스트");
		return "main";
	}
}
