package com.mmcz.Controller;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.websocket.Service;

@RestController
//@RequestMapping("/websocket")
@RequestMapping("/bills")
public class WebSocketController {

	@GetMapping("/pushone")
	public void pushone() throws InterruptedException {
		Service.sendMessage("XXJ", "MMCZ-XXJ");
	}
}