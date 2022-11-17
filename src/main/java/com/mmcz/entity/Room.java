package com.mmcz.entity;

public class Room {
	private String roomid;
	private String roomname;

	public String getRoomid() {
		return roomid;
	}

	public void setRoomid(String roomid) {
		this.roomid = roomid;
	}

	public String getRoomname() {
		return roomname;
	}

	public void setRoomname(String roomname) {
		this.roomname = roomname;
	}

	public int getRoomtype() {
		return roomtype;
	}

	public void setRoomtype(int roomtype) {
		this.roomtype = roomtype;
	}

	public String getRoomcolor() {
		return roomcolor;
	}

	public void setRoomcolor(String roomcolor) {
		this.roomcolor = roomcolor;
	}

	private int roomtype;
	private String roomcolor;
}
