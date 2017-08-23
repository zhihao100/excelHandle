package com.mycomponents.excelhandle.service;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.List;

import org.springframework.stereotype.Service;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.JSONReader;
import com.alibaba.fastjson.JSONWriter;
import com.mycomponents.excelhandle.domain.StudentInfo;
import com.mycomponents.excelhandle.modal.StudentInfoDTO;

@Service
public class JsonHandleService {

	/**
	 * 查询该json文件所有数据
	 * 
	 * @param fileSrc
	 * @return
	 * @throws FileNotFoundException
	 * @throws Exception
	 */
	public List<StudentInfo> readJson(String fileSrc) throws Exception {
		JSONReader reader = new JSONReader(new FileReader(new File(fileSrc)));
		List<StudentInfo> studentInfos = new ArrayList<>();
		reader.startArray();
		while (reader.hasNext()) {
			StudentInfo studentInfo = reader.readObject(StudentInfo.class);
			studentInfos.add(studentInfo);
		}
		reader.endArray();
		reader.close();
		return studentInfos;
	}

	/**
	 * 查询单条json数据
	 * 
	 * @param studentNumber
	 * @return
	 * @throws FileNotFoundException
	 * @throws Exception
	 */
	public StudentInfo readJsonOne(String studentNumber) throws Exception {
		String fileSrc = "src/main/resources/static/studentInfo.json";
		JSONReader reader = new JSONReader(new FileReader(new File(fileSrc)));
		StudentInfo studentInfoResult = null;
		reader.startArray();
		while (reader.hasNext()) {
			StudentInfo studentInfo = reader.readObject(StudentInfo.class);
			if (studentNumber.equals(studentInfo.getStudentNumber())) {
				studentInfoResult = studentInfo;
			}
		}
		reader.endArray();
		reader.close();
		return studentInfoResult;
	}

	public void writeJson(List<StudentInfoDTO> studentInfoDTOList) throws Exception {
		String fileSrc = "src/main/resources/static/studentInfo.json";
		JSONReader reader = new JSONReader(new FileReader(new File(fileSrc)));
		JSONArray arrayData = JSON.parseArray(reader.readObject().toString());
		reader.close();
		JSONWriter writer = new JSONWriter(new FileWriter(fileSrc));
		writer.startArray();
		if (!arrayData.isEmpty()) {
			for (Object data : arrayData) {
				writer.writeValue(data);
			}
		}
		for (int i = 0; i < studentInfoDTOList.size(); i++) {
			writer.writeValue(studentInfoDTOList.get(i));
		}
		writer.endArray();
		writer.close();
	}

	public void updateJson(List<StudentInfoDTO> studentInfoDTOList) throws Exception {
		String fileSrc = "src/main/resources/static/studentInfo.json";
		JSONReader reader = new JSONReader(new FileReader(new File(fileSrc)));
		JSONArray arrayData = JSON.parseArray(reader.readObject().toString());
		reader.close();
		JSONWriter writer = new JSONWriter(new FileWriter(fileSrc));
		writer.startArray();
		for (Object data : arrayData) {
			JSONObject jsonData = JSON.parseObject(JSON.toJSONString(data));
			StudentInfoDTO studentInfoDTO = null;
			for (int i = 0; i < studentInfoDTOList.size(); i++) {
				if (jsonData.get("studentNumber").equals(studentInfoDTOList.get(i).getStudentNumber())) {
					studentInfoDTO = studentInfoDTOList.get(i);
				}
			}
			if (studentInfoDTO != null) {
				writer.writeValue(studentInfoDTO);
			} else {
				writer.writeValue(data);
			}
		}
		writer.endArray();
		writer.close();
	}
}
