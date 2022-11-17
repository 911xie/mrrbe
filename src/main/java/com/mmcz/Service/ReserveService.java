package com.mmcz.Service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.mmcz.Mapper.ReserveMapper;
import com.mmcz.entity.Reserve;

@Service
public class ReserveService {
	@Autowired
	private ReserveMapper reserveMapper;

	public List<Reserve> findAll() {
		return reserveMapper.findAll();
	}

	public Reserve findById(int deptid) {
		return reserveMapper.findById(deptid);
	}

	// 新增数据
	public int save(Reserve reserve) {
		return reserveMapper.save(reserve);
	};

	// 删除数据
	public int delete(int id) {
		return reserveMapper.delete(id);
	};

	// 更新数据
	public int update(Reserve reserve) {
		return reserveMapper.update(reserve);
	};
}
