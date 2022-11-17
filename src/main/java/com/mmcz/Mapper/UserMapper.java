package com.mmcz.Mapper;

import java.util.List;

import org.apache.ibatis.annotations.Delete;
import org.apache.ibatis.annotations.Insert;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.One;
import org.apache.ibatis.annotations.Param;
import org.apache.ibatis.annotations.Result;
import org.apache.ibatis.annotations.Results;
import org.apache.ibatis.annotations.Select;
import org.apache.ibatis.annotations.Update;
import org.apache.ibatis.mapping.FetchType;

import com.mmcz.entity.User;

@Mapper
public interface UserMapper {
	List<User> findAll();

	//这个在UserMapper.xml里配置
	User getUserById(int id);
	
	@Select("select * from user where userid=#{userid}")
	@Results(id = "userMap", value = { 
	@Result(id = true, column = "id", property = "id"),
	@Result(column = "userid", property = "userid"), 
	@Result(column = "username", property = "username"),
	@Result(property = "dept", column = "deptid", 
	one = @One(select = "com.mmcz.Mapper.DeptMapper.findById", fetchType = FetchType.DEFAULT)) })	
	User getUserByUserId(String userid);

	@Select("select * from user")
	@Results(id = "userMapAll", value = { 
	@Result(id = true, column = "id", property = "id"),
	@Result(column = "userid", property = "userid"), 
	@Result(column = "username", property = "username"),
	@Result(property = "dept", column = "deptid", 
	one = @One(select = "com.mmcz.Mapper.DeptMapper.findById", fetchType = FetchType.DEFAULT)) })
	List<User> getAll();
	
	@Select("select * from user where userid=#{userid} and password=#{password}#")
	@Results(id = "findByIdAndPwd", value = { 
	@Result(id = true, column = "id", property = "id"),
	@Result(column = "userid", property = "userid"), 
	@Result(column = "username", property = "username"),
	@Result(property = "dept", column = "deptid", 
	one = @One(select = "com.mmcz.Mapper.DeptMapper.findById", fetchType = FetchType.DEFAULT)) })	
	User findByIdAndPwd(String userid, String password);
	
    // 新增数据
    @Insert("insert into user (type, userid, username, password, deptid, email) values (#{type},#{userid},#{username},#{password},#{deptid},#{email})")
    public int save(User user);
    
    // 删除数据
    @Delete("delete from user where id=#{id}")
    public int delete(int id);

    // 更新数据
    @Update("update user set type=#{type},username=#{username},password=#{password},deptid=#{deptid},email=#{email} where id=#{id}")
    public int update(User user);
    
    // 模糊查询
    @Select("SELECT * FROM user a, dept b where a.deptid=b.deptid and (userid like concat('%',#{text},'%') or username like concat('%',#{text},'%') or email like concat('%',#{text},'%') or b.deptname like concat('%',#{text},'%'))")
	@Results(id = "fuzzyQuery", value = { 
	@Result(id = true, column = "id", property = "id"),
	@Result(column = "userid", property = "userid"), 
	@Result(column = "username", property = "username"),
	@Result(property = "dept", column = "deptid", 
	one = @One(select = "com.mmcz.Mapper.DeptMapper.findById", fetchType = FetchType.DEFAULT)) })	
	List<User> fuzzyQuery(@Param("text") String text);			

}
