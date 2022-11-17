package com.mmcz;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;
import org.springframework.context.annotation.ComponentScan;

//@SpringBootApplication
@ComponentScan("com.*")
@MapperScan("com.mmcz.Mapper")
@SpringBootApplication(exclude = { DataSourceAutoConfiguration.class })
//@SpringBootApplication
public class MrrApplication {

	public static void main(String[] args) {
		SpringApplication.run(MrrApplication.class, args);
	}

}
