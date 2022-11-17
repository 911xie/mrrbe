package com.mmcz.Controller;

import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.mmcz.Service.DeptService;
import com.mmcz.entity.Dept;

@RestController
@RequestMapping("/dept")
public class DeptController {
	@Autowired
	private DeptService deptService;

	@RequestMapping("/findAll")
	public List<Dept> findAll() {
		return deptService.findAll();
	}

	@RequestMapping("/findById")
	public Dept findById(@RequestBody Map<String, String> data) {
		String deptid = data.get("deptid");
		return deptService.findById(Integer.parseInt(deptid));
	}

}
