package egovframework.kss.main.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
public class mainController {

	@RequestMapping("/main.do")
	public String mainPage() {
		return "main";
	}
}
