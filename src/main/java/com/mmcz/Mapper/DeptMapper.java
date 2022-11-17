package com.mmcz.Mapper;

import java.util.List;

import org.apache.ibatis.annotations.Mapper;

import com.mmcz.entity.Dept;

@Mapper
public interface DeptMapper {
	List<Dept> findAll();

	Dept findById(int deptid);
}
