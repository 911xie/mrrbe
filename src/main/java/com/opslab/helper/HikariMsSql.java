package com.opslab.helper;

import java.io.InputStream;
import java.util.Properties;

import com.zaxxer.hikari.HikariConfig;
import com.zaxxer.hikari.HikariDataSource;

public class HikariMsSql {
	private static String driver = null;
	private static String url = null;
	private static String username = null;
	private static String password = null;

	// 在静态代码块中创建数据库连接池
	static {
		try {
			// 加载dbcpconfig.properties配置文件
			InputStream in = HikariMsSql.class.getClassLoader().getResourceAsStream("dbmssql.properties");
			Properties properties = new Properties();
			properties.load(in);
			in.close();

			driver = properties.getProperty("driver");
			url = properties.getProperty("url");
			username = properties.getProperty("username");
			password = properties.getProperty("password");
			System.out.println("HikariDS read config file..." + driver + "," + url + "," + username + "," + password);
		} catch (Exception e) {
			throw new ExceptionInInitializerError(e);
		}
	}
	private static HikariMsSql hikariDS = new HikariMsSql();
	private HikariDataSource dataSource = null;

	// 将构造器设置为private禁止通过new进行实例化
	private HikariMsSql() {
		HikariConfig hikariConfig = new HikariConfig();
		hikariConfig.setJdbcUrl(url);
		hikariConfig.setDriverClassName(driver);
		hikariConfig.setUsername(username);
		hikariConfig.setPassword(password);
		hikariConfig.addDataSourceProperty("cachePrepStmts", "true");
		hikariConfig.addDataSourceProperty("prepStmtCacheSize", "250");
		hikariConfig.addDataSourceProperty("prepStmtCacheSqlLimit", "2048");

		dataSource = new HikariDataSource(hikariConfig);
		System.out.println("Hikari-mssql create...666");
	}

	public static HikariDataSource getDataSource() {
		return hikariDS.dataSource;
	}
}
