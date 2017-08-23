package com.mycomponents.excelhandle.domain;

import java.util.Date;

import com.mycomponents.excelhandle.excel.ExcelColumn;

public class StudentInfo {
	/**
	 * 学号
	 */
	@ExcelColumn(name = "学号", sort = 0)
	private String studentNumber;
	/**
	 * 姓名
	 */
	@ExcelColumn(name = "姓名", sort = 1)
	private String name;
	/**
	 * 性别
	 */
	@ExcelColumn(name = "性别", sort = 2)
	private String gender;
	/**
	 * 手机号码
	 */
	@ExcelColumn(name = "手机号码", sort = 3)
	private String telephoneNumber;
	/**
	 * 学院
	 */
	@ExcelColumn(name = "学院", sort = 4)
	private String institute;
	/**
	 * 年级
	 */
	@ExcelColumn(name = "年级", sort = 5)
	private Integer grade;
	/**
	 * 专业
	 */
	@ExcelColumn(name = "专业", sort = 6)
	private String major;
	/**
	 * 邮箱
	 */
	@ExcelColumn(name = "邮箱", sort = 7)
	private String email;
	/**
	 * 创建时间
	 */
	@ExcelColumn(name = "创建时间", sort = 8)
	private Date createTime;
	/**
	 * 最后一次更新时间
	 */
	@ExcelColumn(name = "最后一次更新时间", sort = 9)
	private Date lastUpdateTime;

	public String getStudentNumber() {
		return studentNumber;
	}

	public void setStudentNumber(String studentNumber) {
		this.studentNumber = studentNumber;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getGender() {
		return gender;
	}

	public void setGender(String gender) {
		this.gender = gender;
	}

	public String getTelephoneNumber() {
		return telephoneNumber;
	}

	public void setTelephoneNumber(String telephoneNumber) {
		this.telephoneNumber = telephoneNumber;
	}

	public String getInstitute() {
		return institute;
	}

	public void setInstitute(String institute) {
		this.institute = institute;
	}

	public Integer getGrade() {
		return grade;
	}

	public void setGrade(Integer grade) {
		this.grade = grade;
	}

	public String getMajor() {
		return major;
	}

	public void setMajor(String major) {
		this.major = major;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public Date getCreateTime() {
		return createTime;
	}

	public void setCreateTime(Date createTime) {
		this.createTime = createTime;
	}

	public Date getLastUpdateTime() {
		return lastUpdateTime;
	}

	public void setLastUpdateTime(Date lastUpdateTime) {
		this.lastUpdateTime = lastUpdateTime;
	}

}
