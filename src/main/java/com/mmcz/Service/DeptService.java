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

	public Dept findById(int deptid) {
		return deptMapper.findById(deptid);
	}
}
