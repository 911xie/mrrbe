package com.common.Interceptor;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.InterceptorRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

@Configuration
public class InterceptorConfig implements WebMvcConfigurer {
	@Override
	public void addInterceptors(InterceptorRegistry registry) {
		System.out.println("addInterceptors...");
		registry.addInterceptor(jwtInterceptor()).addPathPatterns("/user/**", "/reserve/**", "/dept/**") // 拦截所有请求，通过判断token是否合法
				.excludePathPatterns("/user/login", "/file/upload", "/bills");// 需要放行的方法, "/api"
	}

	@Bean
	public JwtInterceptor jwtInterceptor() {
		return new JwtInterceptor();
	}
}
