package com.opslab.helper;

import java.util.Calendar;
import java.util.Date;
import java.util.List;

import com.auth0.jwt.JWT;
import com.auth0.jwt.algorithms.Algorithm;
import com.auth0.jwt.interfaces.Claim;
import com.auth0.jwt.interfaces.DecodedJWT;
import com.auth0.jwt.interfaces.JWTVerifier;

public class JWTUtil {
	public static final String suffix = "JwTkEy";

	/**
	 * 获取令牌
	 *
	 * @param userId   用户Id
	 * @param username 用户名
	 * @param userNick 用户昵称
	 * @return 令牌
	 */
	public static String generateToken(String userId, String username, String userNick) {
		Calendar calendar = Calendar.getInstance();
		Date issueDate = calendar.getTime();
		calendar.add(Calendar.MINUTE, 60);
		Date expireDate = calendar.getTime();
		return JWT.create()
				// 签发对象
				.withAudience(userId)
				// 发行时间
				.withIssuedAt(issueDate)
				// 过期时间
				.withExpiresAt(expireDate)
				// 载荷
				.withClaim("username", username)
				// 载荷
				.withClaim("userNick", userNick)
				// 签名
				// 使用HMAC SHA256签名算法。
				// HMAC（全称Hash-based Message Authentication Code），基于Hash函数和密钥的消息认证码。
				// 密钥是userId + suffix
				.sign(Algorithm.HMAC256(userId + suffix));
	}

	/**
	 * 检验令牌
	 * 
	 * @param token        令牌
	 * @param secretPrefix 密钥前缀
	 */
	public static void verifyToken(String token, String secretPrefix) {
		JWTVerifier verifier = JWT.require(Algorithm.HMAC256(secretPrefix + suffix)).build();
		DecodedJWT decodedJWT = verifier.verify(token);
	}

	/**
	 * 判断 token 是否过期
	 * 
	 * @param token 令牌
	 * @return 是否过期，过期true，未过false
	 */
	public static boolean isExpire(String token) {
		DecodedJWT jwt = JWT.decode(token);
		// 如果token的过期时间小于当前时间，则表示已过期，为true
		return jwt.getExpiresAt().getTime() < System.currentTimeMillis();
	}

	/**
	 * 获取签发对象
	 *
	 * @param token 令牌
	 * @return 签发对象
	 */
	public static String getAudience(String token) {
		List<String> audiences = JWT.decode(token).getAudience();
		String audience = audiences.get(0);
		return audience;
	}

	/**
	 * 通过载荷名称获取载荷值
	 *
	 * @param token     令牌
	 * @param claimName 载荷名称
	 * @return 载荷值
	 */
	public static Claim getClainByName(String token, String claimName) {
		return JWT.decode(token).getClaim(claimName);
	}

	/**
	 * 测试
	 */
	public static void main(String[] args) throws InterruptedException {
		String token = JWTUtil.generateToken("100001", "juyoujing", "橘右京");
		System.out.println("令牌：" + token);

		JWTUtil.verifyToken(token, "100001");

		String audience = JWTUtil.getAudience(token);
		System.out.println("签发对象：" + audience);

		Claim claim = JWTUtil.getClainByName(token, "userNick");
		System.out.println("载荷：" + claim.asString());

	}
}
