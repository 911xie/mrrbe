package com.mmcz.entity;

public class Dept {

	private Integer deptid;
	private Integer parentid;

	public Integer getParentid() {
		return parentid;
	}

	public void setParentid(Integer parentid) {
		this.parentid = parentid;
	}

	public Integer getLevel() {
		return level;
	}

	public void setLevel(Integer level) {
		this.level = level;
	}

	private Integer level;
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

	public Dept() {
	}

	@Override
	public String toString() {
		return "Dept{" + "deptid=" + deptid + ", parentid=" + parentid + ", level=" + level + ", deptname='" + deptname
				+ '\'' + '}';
	}
}