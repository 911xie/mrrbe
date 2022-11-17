package com.opslab.helper;

import java.util.Hashtable;

import javax.naming.AuthenticationException;
import javax.naming.Context;
import javax.naming.NamingEnumeration;
import javax.naming.NamingException;
import javax.naming.directory.Attribute;
import javax.naming.directory.Attributes;
import javax.naming.directory.SearchControls;
import javax.naming.directory.SearchResult;
import javax.naming.ldap.Control;
import javax.naming.ldap.InitialLdapContext;
import javax.naming.ldap.LdapContext;

public class LDAP {
	private String URL = "LDAP://MCZADSRV01.MMCZ.MEKJPN.NGNET/";
	private String FACTORY = "com.sun.jndi.ldap.LdapCtxFactory";
	private String BASEDN = "DC=MMCZ,DC=MEKJPN,DC=NGNET";// "DC=mmcz,DC=mekjpn,DC=ngnet";
	private LdapContext ctx = null;
	private Hashtable env = null;
	private Control[] connCtls = null;

	private void LDAPConnect() {
		env = new Hashtable();
		env.put(Context.INITIAL_CONTEXT_FACTORY, FACTORY);
		env.put(Context.PROVIDER_URL, URL + BASEDN);
		env.put(Context.SECURITY_AUTHENTICATION, "simple");
		env.put(Context.SECURITY_PRINCIPAL, "MMCZ_MANAGEUSER");
		env.put(Context.SECURITY_CREDENTIALS, "#Zixiang#.Pwr2022");

		// 此处若不指定用户名和密码,则自动转换为匿名登录
		try {
			ctx = new InitialLdapContext(env, connCtls);
			SearchControls searchCtls = new SearchControls();
			searchCtls.setSearchScope(SearchControls.SUBTREE_SCOPE);
			searchCtls.setCountLimit(1000);
			String searchFilter = "sAMAccountName=XIE*";
			// String searchFilter = "sAMAccountName=*";
			// String searchFilter = "physicalDeliveryOfficeName=*";
			String searchBase = BASEDN;
			String returnedAtts[] = { "*", "createTimeStamp", "modifyTimeStamp", "structuralObjectClass",
					"subSchemaSubEntry" };
			searchCtls.setReturningAttributes(returnedAtts);
			NamingEnumeration<SearchResult> answer = ctx.search("", searchFilter, searchCtls);
			System.out.println(answer);

			int totalResults = 0;// Specify the attributes to return
			int rows = 0;
			String allName = "";
			while (answer.hasMoreElements()) {// 遍历结果集
				SearchResult sr = (SearchResult) answer.next();// 得到符合搜索条件的DN
				System.out.println(++rows + "************************************************");

				String dn = sr.getName();
				System.out.println(dn);
				String match = dn.split("CN=")[1].split(",")[0];// 返回格式一般是CN=ptyh,OU=专卖
				System.out.println(match);
				Attributes Attrs = sr.getAttributes();// 得到符合条件的属性集
				if (Attrs != null) {
					try {
						for (NamingEnumeration ne = Attrs.getAll(); ne.hasMore();) {
							Attribute Attr = (Attribute) ne.next();// 得到下一个属性
							String key = Attr.getID().toString();
							System.out.println(" AttributeID=属性名：" + Attr.getID().toString());
							// 读取属性值
							for (NamingEnumeration e = Attr.getAll(); e.hasMore(); totalResults++) {
								String AttrValue = e.next().toString();
								System.out.println("    AttributeValues=属性值：" + AttrValue);
							}
							System.out.println("    ---------------");
						}
					} catch (NamingException e) {
						System.err.println("Throw Exception : " + e);
					}
				} // if
					// }
			} // while
			ctx.close();

			System.out.println("InitialLdapContext success!" + allName);
		} catch (NamingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private String getUserDN(String email) {
		String userDN = "";
		LDAPConnect();
		try {
			String filters = "(&(&(objectCategory=person)(objectClass=user))(sAMAccountName=elbert.chenh))";
			String[] returnedAtts = { "distinguishedName", "userAccountControl", "displayName", "employeeID" };
			SearchControls constraints = new SearchControls();
			constraints.setSearchScope(SearchControls.SUBTREE_SCOPE);
			if (returnedAtts != null && returnedAtts.length > 0) {
				constraints.setReturningAttributes(returnedAtts);
			}

			NamingEnumeration en = ctx.search(BASEDN, filters, constraints);
			if (en == null) {
				System.out.println("Have no NamingEnumeration.");
			}

			if (!en.hasMoreElements()) {
				System.out.println("Have no element.");
			} else {

				while (en != null && en.hasMoreElements()) {
					Object obj = en.nextElement();
					if (obj instanceof SearchResult) {
						SearchResult si = (SearchResult) obj;
						Attributes userInfo = si.getAttributes();
						userDN += userInfo.toString();
						userDN += "," + BASEDN;
					} else {
						System.out.println(obj.toString());
					}
					System.out.println(userDN);
				}

			}
		} catch (Exception e) {
			System.out.println("Exception in search():" + e);
		}
		return userDN;
	}

	public boolean authenricate(String ID, String password) {
		boolean valide = false;
		String userDN = getUserDN(ID);
		try {
			ctx.addToEnvironment(Context.SECURITY_PRINCIPAL, userDN);
			ctx.addToEnvironment(Context.SECURITY_CREDENTIALS, password);
			ctx.reconnect(connCtls);
			System.out.println(userDN + " is authenticated");
			valide = true;
		} catch (AuthenticationException e) {
			System.out.println(userDN + " is not authenticated");
			System.out.println(e.toString());
			valide = false;
		} catch (NamingException e) {
			System.out.println(userDN + " is not authenticated");
			valide = false;
		}
		return valide;
	}

	public static int gi() {
		int b = 10;
		try {
			System.out.println(b / 0);
			b = 99;
		} catch (ArithmeticException e) {
			b = 66;
			return b;
		} finally {
			b = 20;
		}
		return b;
	}

	public static void main(String[] args) throws Exception {
		LDAP ldap = new LDAP();
		ldap.LDAPConnect();
		// System.out.println("resultttttt=" + LDAP.gi());

	}

}