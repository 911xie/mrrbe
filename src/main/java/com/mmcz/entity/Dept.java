package com.mmcz.entity;

public class Dept {

	private Integer deptid;
	private String deptname;

	public Integer getDeptid() {
		return deptid;
	}

	public void setDeptid(Integer deptid) {
		this.deptid = deptid;
	}

	public String getDeptname() {
		return deptname;
	}

	public void setDeptname(String deptname) {
		this.deptname = deptname;
	}

	public Dept(int deptid, String deptname) {
		this.deptid = deptid;
		this.deptname = deptname;
	}

	@Override
	public String toString() {
		return "Dept{" + "deptid=" + deptid + ", deptname='" + deptname + '\'' + '}';
	}
}