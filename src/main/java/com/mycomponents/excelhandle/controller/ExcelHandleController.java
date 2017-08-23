package com.mycomponents.excelhandle.controller;

import java.io.InputStream;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;
import org.springframework.web.servlet.view.RedirectView;

import com.mycomponents.excelhandle.domain.StudentInfo;
import com.mycomponents.excelhandle.service.ExcelHandleService;
import com.mycomponents.excelhandle.service.JsonHandleService;

@RestController
public class ExcelHandleController {
	@Autowired
	private ExcelHandleService excelHandleService;
	@Autowired
	private JsonHandleService jsonHandleService;

	@RequestMapping("/")
	public ModelAndView toIndex() {
		return new ModelAndView(new RedirectView("templates/index.html"));
	}

	/**
	 * 导出全部供查阅
	 * 
	 * @param infoList
	 * @param response
	 * @throws Exception
	 */
	@RequestMapping(value = "/exportExcel", method = RequestMethod.GET)
	public void exportExcel(HttpServletResponse response) throws Exception {
		String fileSrc = "src/main/resources/static/studentInfo.json";
		// 获取json测试数据
		List<StudentInfo> infoList = jsonHandleService.readJson(fileSrc);
		excelHandleService.exportExcel(infoList, response);
	}

	/**
	 * 导出供导入使用的模板
	 * 
	 * @param infoList
	 * @param response
	 * @throws Exception
	 */
	@RequestMapping(value = "/getImportTemplateExcel", method = RequestMethod.GET)
	public void getImportTemplateExcel(HttpServletResponse response) throws Exception {
		String fileSrc = "src/main/resources/static/studentInfo.json";
		// 获取json测试数据
		List<StudentInfo> infoList = jsonHandleService.readJson(fileSrc);
		excelHandleService.getImportTemplateExcel(infoList, response);
	}

	/**
	 * 导入
	 * 
	 * @param input
	 * @param fileName
	 * @throws Exception
	 */
	@RequestMapping(value = "/importExcel", method = RequestMethod.POST)
	public String importExcel(@RequestBody MultipartFile file) throws Exception {
		String fileName = file.getOriginalFilename();
		InputStream input = file.getInputStream();
		return excelHandleService.importExcel(input, fileName);
	}
}
