package com.mmcz.Mapper;

import java.util.List;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Select;

import com.mmcz.entity.Room;

@Mapper
public interface RoomMapper {
	List<Room> findAll();

	@Select("select * from room where roomid=#{roomid}")
	Room findById(int roomid);
}
