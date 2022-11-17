package com.mmcz.entity;

import javax.validation.constraints.Min;
import javax.validation.constraints.NotEmpty;

import lombok.Data;

@Data
public class User {
	private int id;

	@Min(value = 0, message = "用户类型不允许为负数")
	private int type;

	@NotEmpty(message = "用户id不允许为空")
	private String userid;
	private String username;
	private String password;
	private String email;
	private int deptid;

	public int getDeptid() {
		return deptid;
	}

	public void setDeptid(int deptid) {
		this.deptid = deptid;
	}

	private Dept dept;

	public Dept getDept() {
		return dept;
	}

	public void setDept(Dept dept) {
		this.dept = dept;
	}

	public int getType() {
		return type;
	}

	public void setType(int type) {
		this.type = type;
	}

	public String getUserid() {
		return userid;
	}

	public void setUserid(String userid) {
		this.userid = userid;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public User(int id, int type, String userid, String username, String password, String email, Dept dept) {
		this.id = id;
		this.type = type;
		this.userid = userid;
		this.username = username;
		this.password = password;
		this.dept = dept;
		this.email = email;
	}

	public User() {
	}

	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}

	public String getUsername() {
		return username;
	}

	public void setUsername(String username) {
		this.username = username;
	}

	public String getPassword() {
		return password;
	}

	public void setPassword(String password) {
		this.password = password;
	}

	@Override
	public String toString() {
		return "User{" + "id=" + id + ", type=" + type + ", userid=" + userid + ", username='" + username + '\''
				+ ", password='" + password + '\'' + ", deptid=" + dept.getDeptid() + ", email='" + email + '\'' + '}';
	}
}
