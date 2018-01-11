package com.tops.common.util;

import java.io.File;

public class FileUtil {
	private FileUtil(){}

	/**
	 * @param file pathname
	 * @return boolean
	 */
	public static boolean fileExist(String file){
		File f = new File(file);
		return f.exists();
	}
	
	/**
	 * 
	 * @param file pathname
	 * @return cast file size
	 */
	public static String getFileSize(String file){
		long length = new File(file).length();
		return castFileSize(length);
	}
	
	/**
	 * @param length file size
	 * @return cast file size
	 */
	private static String castFileSize(long length){
		if(length < 1024){
			return length + "B";
		}else if(length<1048576){
			return length / 1024 + "K";
		}else if(length < 1073741824){
			return length / 1048576 + "M";
		}else{
			return length / 1073741824 + "G";
		}
	}
	
	public static void main(String[] args) {
		System.out.println(fileExist("D:/test"));
		System.out.println(getFileSize("D:/software/我的文档/QQ/1258123512/FileRecv/灾毁报送表系统及已发现问题.docx"));
	}
}
