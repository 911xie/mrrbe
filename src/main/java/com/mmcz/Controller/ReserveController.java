package com.mmcz.Controller;

import java.io.UnsupportedEncodingException;
import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.alibaba.fastjson.JSON;
import com.common.response.ResultBody;
import com.common.response.ResultBodyUtil;
import com.mmcz.Service.ReserveService;
import com.mmcz.entity.Reserve;
import com.opslab.helper.CharsetHepler;

@RestController
@RequestMapping("/reserve")
public class ReserveController {
	@Autowired
	private ReserveService reserveService;

	@RequestMapping("/findAll")
	public ResultBody findAll() {
		List<Reserve> reserveList = reserveService.findAll();
		ResultBody result = ResultBodyUtil.success(reserveList);
		return result;
	}

	@RequestMapping("/findById")
	public ResultBody findById(@RequestBody Map<String, String> data) {
		String reserveid = data.get("id");
		Reserve reserve = reserveService.findById(Integer.parseInt(reserveid));
		return ResultBodyUtil.success(reserve);
	}

	@RequestMapping("/add")
	public ResultBody add(@RequestBody @Validated Reserve reserve) throws UnsupportedEncodingException {
		System.out.println("addUser-in " + CharsetHepler.toGBK(JSON.toJSONString(reserve)));
		reserveService.save(reserve);
		ResultBody result = ResultBodyUtil.success();
		return result;
	}

	@RequestMapping("/update")
	public ResultBody update(@RequestBody Reserve reserve) throws UnsupportedEncodingException {
		System.out.println("update-in " + CharsetHepler.toGBK(JSON.toJSONString(reserve)));
		reserveService.update(reserve);
		return ResultBodyUtil.success();
	}

	@RequestMapping("/delete")
	public ResultBody delete(@RequestBody Map<String, String> data) {
		Integer id = Integer.parseInt(data.get("id"));
		reserveService.delete(id);
		System.out.println("delete " + id);
		return ResultBodyUtil.success();
	}
}
