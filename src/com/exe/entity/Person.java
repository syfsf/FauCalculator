package com.exe.entity;

public class Person {
	private int companyCode;
	private String endDate;	// 统计截止日期
	private String name;
	private double sex; // 0代表女性，1代表男性
	private double age;
	private double education; // 学历等级
	private double workMonth; // 工作年限，以月数计

	public String getEndDate() {
		return endDate;
	}

	public void setEndDate(String endDate) {
		this.endDate = endDate;
	}

	public int getCompanyCode() {
		return companyCode;
	}

	public void setCompanyCode(int companyCode) {
		this.companyCode = companyCode;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public double getSex() {
		return sex;
	}

	public void setSex(double sex) {
		this.sex = sex;
	}

	public double getAge() {
		return age;
	}

	public void setAge(double age) {
		this.age = age;
	}

	public double getEducation() {
		return education;
	}

	public void setEducation(double education) {
		this.education = education;
	}

	public double getWorkMonth() {
		return workMonth;
	}

	public void setWorkMonth(double workMonth) {
		this.workMonth = workMonth;
	}

	@Override
	public String toString() {
		return "Person [companyCode=" + companyCode + ", endDate=" + endDate + ", name=" + name + ", sex=" + sex
				+ ", age=" + age + ", education=" + education + ", workMonth=" + workMonth + "]";
	}

}
