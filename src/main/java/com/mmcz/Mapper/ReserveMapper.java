package com.mmcz.Mapper;

import java.sql.Timestamp;
import java.util.List;

import org.apache.ibatis.annotations.Delete;
import org.apache.ibatis.annotations.Insert;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.One;
import org.apache.ibatis.annotations.Result;
import org.apache.ibatis.annotations.Results;
import org.apache.ibatis.annotations.Select;
import org.apache.ibatis.annotations.Update;
import org.apache.ibatis.mapping.FetchType;

import com.mmcz.entity.Reserve;
import com.mmcz.entity.User;

@Mapper
public interface ReserveMapper {
	@Select("select * from reserve")
	@Results(id = "ReserveFindAll", value = { 
	@Result(id = true, column = "id", property = "id"),
	@Result(column = "title", property = "title"), 
	@Result(column = "content", property = "content"),
	@Result(column = "begin", property = "begin"), 
	@Result(column = "create", property = "create"),
	@Result(property = "user", column = "userid", 
	one = @One(select = "com.mmcz.Mapper.UserMapper.getUserByUserId", fetchType = FetchType.DEFAULT)), 
	@Result(property = "room", column = "roomid", 
	one = @One(select = "com.mmcz.Mapper.RoomMapper.findById", fetchType = FetchType.DEFAULT))})
	List<Reserve> findAll();

	@Select("select * from reserve where id=#{reserveid}")
	@Results(id = "ReserveFindById", value = { 
	@Result(id = true, column = "id", property = "id"),
	@Result(column = "title", property = "title"), 
	@Result(column = "content", property = "content"),
	@Result(column = "begin", property = "begin"), 
	@Result(column = "create", property = "create"),
	@Result(property = "user", column = "userid", 
	one = @One(select = "com.mmcz.Mapper.UserMapper.getUserByUserId", fetchType = FetchType.DEFAULT)), 
	@Result(property = "room", column = "roomid", 
	one = @One(select = "com.mmcz.Mapper.RoomMapper.findById", fetchType = FetchType.DEFAULT))})
	Reserve findById(int reserveid);
	
    // 新增数据
    @Insert("insert into reserve(userid,title,content,start,end,roomid) values(#{userid},#{title},#{content},#{start},#{end}, #{roomid})")
    public int save(Reserve reserve);
    
    // 删除数据
    @Delete("delete from reserve where id=#{id}")
    public int delete(int id);

    // 更新数据
    @Update("update reserve set userid=#{userid},title=#{title},content=#{content},start=#{start},end=#{end},roomid=#{roomid} where id=#{id}")
    public int update(Reserve reserve);
}
