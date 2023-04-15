package com.mmcz.Service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.mmcz.Mapper.DeptMapper;
import com.mmcz.entity.Dept;

@Service
public class DeptService {
	@Autowired
	private DeptMapper deptMapper;

	public List<Dept> findAll() {
		return deptMapper.findAll();
	}

	public int update(Dept dept) {
		return deptMapper.update(dept);
	};

	public int add(Dept dept) {
		deptMapper.add(dept);
		int maxid = dept.getDeptid();
		return maxid;
	};

	public int delete(List<String> ids) {
		return deptMapper.delete(ids);
	};

	public Dept findById(int deptid) {
		return deptMapper.findById(deptid);
	}
}
