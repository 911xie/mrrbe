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
		registry.addInterceptor(jwtInterceptor()).addPathPatterns("/user/**", "/reserve/**") // 拦截所有请求，通过判断 token是否合法
				.excludePathPatterns("/user/login", "/file/upload", "/bills", "/api");// 需要放行的方法
	}

	@Bean
	public JwtInterceptor jwtInterceptor() {
		return new JwtInterceptor();
	}
}
