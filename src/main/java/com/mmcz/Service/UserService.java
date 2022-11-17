package com.mmcz.Service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.mmcz.Mapper.UserMapper;
import com.mmcz.entity.User;

@Service
public class UserService {
	@Autowired
	private UserMapper userMapper;

	public List<User> findAll() {
		return userMapper.findAll();
	}

	public List<User> getAll() {
		return userMapper.getAll();
	}

	public User getUserById(int id) {
		return userMapper.getUserById(id);
	}

	public User getUserByUserId(String userid) {
		return userMapper.getUserByUserId(userid);
	}

	public User findByIdAndPwd(String userid, String password) {
		return userMapper.findByIdAndPwd(userid, password);
	}

	// 新增数据
	public int save(User user) {
		return userMapper.save(user);
	};

	// 删除数据
	public int delete(int id) {
		return userMapper.delete(id);
	};

	// 更新数据
	public int update(User user) {
		return userMapper.update(user);
	};

	// 模糊查询
	public List<User> fuzzyQuery(String text) {
		return userMapper.fuzzyQuery(text);
	};
}
