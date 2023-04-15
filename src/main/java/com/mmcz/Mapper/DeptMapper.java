package com.mmcz.Mapper;

import java.util.List;

import org.apache.ibatis.annotations.Delete;
import org.apache.ibatis.annotations.Insert;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Options;
import org.apache.ibatis.annotations.Update;

import com.mmcz.entity.Dept;

@Mapper
public interface DeptMapper {
	List<Dept> findAll();

	// 更新数据
	@Update("update dept set deptname=#{deptname} where deptid=#{deptid}")
	public int update(Dept dept);

	// 添加数据
	@Insert("insert into dept(parentid,deptname) values(#{parentid},#{deptname})")
	@Options(useGeneratedKeys = true, keyProperty = "deptid", keyColumn = "deptid")
	public int add(Dept dept);

	@Delete("<script> DELETE FROM dept WHERE deptid in <foreach collection='ids' item='id' open='(' separator=',' close=')'>#{id}</foreach> </script>")
//    int deleteByIdBatch(List<String> ids);
//	@Delete("delete from dept where deptid in (#{ids})")
	public int delete(List<String> ids);

	Dept findById(int deptid);
}
