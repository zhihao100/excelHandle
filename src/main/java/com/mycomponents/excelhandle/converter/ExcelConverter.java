package com.mycomponents.excelhandle.converter;

import java.util.Map;

import com.mycomponents.excelhandle.modal.StudentInfoDTO;

public class ExcelConverter {
	/**
	 * 将从excel读取的map数据转换成StudentInfoDTO模型
	 * 
	 * @param map
	 * @return
	 */
	public static StudentInfoDTO convertToStudentInfoDTO(Map<String, Object> map) {
		StudentInfoDTO dto = new StudentInfoDTO();
		if (map != null) {
			String stuNo = map.get("studentNumber").toString();
			if (stuNo.matches("^[0-9]+(.0)$")) {
				stuNo.substring(0, stuNo.length() - 2);
			}
			if (stuNo.matches("^[0-9]+(.)\\d{9}+(E9)$")) {
				String stuNoArray = stuNo.substring(0, 11);
				int pointPos = stuNoArray.indexOf(".");
				stuNo = stuNoArray.substring(0, pointPos) + stuNoArray.substring(pointPos + 1);
			}
			dto.setStudentNumber(stuNo);
			dto.setName(map.get("name").toString());
			dto.setGender(map.get("gender").toString());
			dto.setInstitute(map.get("institute").toString());
			dto.setMajor(map.get("major").toString());
			String grade = map.get("grade").toString();
			if (grade.matches("^[20]+([0-9]{2})+(.0)$")) {
				grade = grade.substring(0, grade.length() - 2);
			}
			dto.setGrade(Integer.valueOf(grade));
			String telNumber = map.get("telephoneNumber").toString();
			if (telNumber.matches("^(?:\\+86)?1+(.)\\d{10}+(E10)$")) {
				String telNumberArray = telNumber.substring(0, 12);
				int pointPos = telNumberArray.indexOf(".");
				telNumber = telNumberArray.substring(0, pointPos) + telNumberArray.substring(pointPos + 1);
			}
			dto.setTelephoneNumber(telNumber);
			dto.setEmail(map.get("email").toString());
		}
		return dto;
	}
}
