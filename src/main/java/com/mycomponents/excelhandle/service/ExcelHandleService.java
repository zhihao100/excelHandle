package com.mycomponents.excelhandle.service;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.mycomponents.excelhandle.converter.ExcelConverter;
import com.mycomponents.excelhandle.domain.StudentInfo;
import com.mycomponents.excelhandle.excel.ExcelUtils;
import com.mycomponents.excelhandle.modal.StudentInfoDTO;
import com.mycomponents.excelhandle.modal.StudentInfoVO;

@Service
public class ExcelHandleService {
	@Autowired
	private JsonHandleService jsonHandleService;

	/**
	 * 导出全部供查阅
	 * 
	 * @param infoList
	 * @param response
	 * @throws Exception
	 */
	public void exportExcel(List<StudentInfo> infoList, HttpServletResponse response) throws Exception {
		HSSFWorkbook wb = ExcelUtils.exportReflectInfo(StudentInfo.class, infoList, "学生信息");
		String fileName = "学生信息一览表.xls";
		response.reset();
		response.setContentType("application/vnd.ms-excel;charset=utf-8");
		response.setHeader("Content-Disposition",
				"attachment;filename=" + new String(fileName.getBytes(), "ISO8859-1"));
		response.setCharacterEncoding("utf-8");
		OutputStream os = response.getOutputStream();
		wb.write(os);
		os.flush();
		os.close();
		wb.close();
	}

	/**
	 * 导出供导入使用的模板
	 * 
	 * @param infoList
	 * @param response
	 * @throws Exception
	 */
	public void getImportTemplateExcel(List<StudentInfo> infoList, HttpServletResponse response) throws Exception {
		HSSFWorkbook wb = ExcelUtils.exportReflectInfo(StudentInfoVO.class, infoList, "学生信息模板");
		String fileName = "学生信息表模板.xls";
		response.reset();
		response.setContentType("application/vnd.ms-excel;charset=utf-8");
		response.setHeader("Content-Disposition",
				"attachment;filename=" + new String(fileName.getBytes(), "ISO8859-1"));
		response.setCharacterEncoding("utf-8");
		OutputStream os = response.getOutputStream();
		wb.write(os);
		os.flush();
		os.close();
		wb.close();
	}

	/**
	 * 导入
	 * 
	 * @param input
	 * @param fileName
	 * @throws Exception
	 */
	public String importExcel(InputStream input, String fileName) throws Exception {
		List<Map<String, Object>> studentInfoVOList = ExcelUtils.importReflectInfo(StudentInfoVO.class, input,
				fileName);
		List<StudentInfoDTO> studentInfoDTOList = new ArrayList<>();
		// 设置校验结果信息返回给前台,并读取正确数据到studentInfoDTOList
		StringBuilder message = new StringBuilder();
		int m = 1;
		message.append("第");
		for (Map<String, Object> map : studentInfoVOList) {
			m++;
			boolean isCorrect = validateStudentInfoVO(map);
			if (isCorrect) {
				studentInfoDTOList.add(ExcelConverter.convertToStudentInfoDTO(map));
			} else {
				message.append(m + "、");
			}
		}
		message.deleteCharAt(message.length() - 1);
		if (message.length() != 0) {
			message.append("行数据校验不通过，学号只能是数字，手机号为1开头的11位数字，年级以20开头的4位数字，邮件为带有@的正确格式字符，性别为“男”或“女”，其余项为中文");
		}
		// 以学号为准判断是执行新增还是更新操作
		List<StudentInfoDTO> writeList = new ArrayList<>();
		List<StudentInfoDTO> updateList = new ArrayList<>();
		for (StudentInfoDTO studentInfoDTO : studentInfoDTOList) {
			String stuNo = studentInfoDTO.getStudentNumber();
			if (findExist(stuNo) != null) {
				studentInfoDTO.setCreateTime(findExist(stuNo).getCreateTime());
				studentInfoDTO.setLastUpdateTime(new Date());
				updateList.add(studentInfoDTO);
			} else {
				studentInfoDTO.setCreateTime(new Date());
				studentInfoDTO.setLastUpdateTime(new Date());
				writeList.add(studentInfoDTO);
			}
		}
		if (!updateList.isEmpty()) {
			jsonHandleService.updateJson(updateList);
		}
		if (!writeList.isEmpty()) {
			jsonHandleService.writeJson(writeList);
		}
		return message.toString();
	}

	// 查询json测试文件该学生是否已存在
	private StudentInfo findExist(String stuNo) throws Exception {
		StudentInfo studentInfo = jsonHandleService.readJsonOne(stuNo);
		return studentInfo;
	}

	private Boolean validateStudentInfoVO(Map<String, Object> map) {
		boolean isCorrect = true;
		for (Map.Entry<String, Object> entry : map.entrySet()) {
			boolean valueEmpty = entry.getValue() == null;
			if (valueEmpty) {
				isCorrect = false;
			}

			boolean valueNotEmpty = entry.getValue() != null;
			String value = "";
			String key = "";
			if (valueNotEmpty) {
				value = entry.getValue().toString();
				key = entry.getKey();
			}
			if (valueNotEmpty && ("studentNumber".equals(key))) {
				if (!((value.matches("^[0-9]+$") || value.matches("^[0-9]+(.0)$")
						|| value.matches("^[0-9]+(.)\\d{9}+(E9)$")))) {
					isCorrect = false;
				}
			} else if (valueNotEmpty && "telephoneNumber".equals(key)) {
				if (!((value.matches("^(?:\\+86)?1\\d{10}$") || value.matches("^(?:\\+86)?1+(.)\\d{10}+(E10)$")))) {
					isCorrect = false;
				}
			} else if (valueNotEmpty && "grade".equals(key)) {
				if (!(value.matches("^[20]+([0-9]{2})$") || value.matches("^[20]+([0-9]{2})+(.0)$"))) {
					isCorrect = false;
				}
			} else if (valueNotEmpty && "email".equals(key)) {
				if (!value.matches("^\\w+([-+.]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*$")) {
					isCorrect = false;
				}
			} else if (valueNotEmpty && "gender".equals(key)) {
				if (!("男".equals(value) || "女".equals(value))) {
					isCorrect = false;
				}
			} else if (valueNotEmpty) {
				if (!value.matches("^[\u4E00-\u9FA5]+$")) {
					isCorrect = false;
				}
			}
		}

		return isCorrect;
	}

}
