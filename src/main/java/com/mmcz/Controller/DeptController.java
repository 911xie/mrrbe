package com.mmcz.Controller;

import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.common.response.ResultBody;
import com.common.response.ResultBodyUtil;
import com.mmcz.Service.DeptService;
import com.mmcz.entity.Dept;

@RestController
@RequestMapping("/dept")
public class DeptController {
	@Autowired
	private DeptService deptService;

	@RequestMapping("/findAll")
//	public List<Dept> findAll() {
//		return deptService.findAll();
//	}
	public ResultBody findAll() {
		List<Dept> deptList = deptService.findAll();
		ResultBody result = ResultBodyUtil.success(deptList);
		return result;
	}

	@RequestMapping("/update")
	// public ResultBody update(@RequestBody Map<String, String> data) {
	public ResultBody update(@RequestBody Dept dept) {
		System.out.println(dept.toString());
		deptService.update(dept);
		ResultBody result = ResultBodyUtil.success();
		return result;
	}

	@RequestMapping("/add")
	// public ResultBody update(@RequestBody Map<String, String> data) {
	public ResultBody add(@RequestBody Dept dept) {
		System.out.println(dept.toString());
		int maxid = deptService.add(dept);
		System.out.println("maxid:" + maxid);
		Dept rtnDept = new Dept();
		rtnDept.setDeptid(maxid);
		rtnDept.setParentid(dept.getParentid());
		rtnDept.setDeptname(dept.getDeptname());
		ResultBody result = ResultBodyUtil.success(rtnDept);
		return result;
	}

	@RequestMapping("/delete")
	// public ResultBody update(@RequestBody Map<String, String> data) {
	public ResultBody delete(@RequestBody Map<String, List<String>> data) {
		// String ids = data.get("ids");

		List<String> idlist = data.get("ids");
		System.out.println(idlist);
		deptService.delete(idlist);
		ResultBody result = ResultBodyUtil.success();
		return result;
	}

	@RequestMapping("/findById")
	public Dept findById(@RequestBody Map<String, String> data) {
		String deptid = data.get("deptid");
		return deptService.findById(Integer.parseInt(deptid));
	}

}
