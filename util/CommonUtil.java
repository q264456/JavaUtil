package com.tops.common.util;

import java.util.UUID;

public class CommonUtil {
	/**
	 * @return uuid
	 */
	public static String getUUID(){
		return UUID.randomUUID().toString().replace("-", "");
	}
	
	
	
	
	
	
	
	public static void main(String[] args) {
		System.out.println(getUUID());
	}
}
