/**
 * 
 */
package cn.com.jcom.util;

import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;

import cn.com.jcom.model.Common;


/**
 * <P>
 * 类说明：<br/>
 * 说明写这儿
 *
 * 
 * </p>
 * 创建者：@author ( admin)<br/>
 * 创建时间： 2016-1-25<br/>
 * 
 */

/**
 * <P>
 * 类说明：<br/>
 * 说明写这儿
 *
 * 
 * </p>
 * 创建者：@author ( admin)<br/>
 * 创建时间： 2016-1-25<br/>
 */
public class Util {
	
	 /**
     * get postfix of the path
     * @param path
     * @return
     */
    public static String getPostfix(String path) {
        if (path == null || Common.EMPTY.equals(path.trim())) {
            return Common.EMPTY;
        }
        if (path.contains(Common.POINT)) {
            return path.substring(path.lastIndexOf(Common.POINT) + 1, path.length());
        }
        return Common.EMPTY;
    }
    
    public static String getCurrentJarPath(){
    	String path = Util.class.getProtectionDomain().getCodeSource().getLocation().getPath();
		path = path.substring(1);
		path = path.substring(0,path.lastIndexOf("/") );
		path = path.substring(0,path.lastIndexOf("/") );
		path = path.substring(0,path.lastIndexOf("/") );
		try {
			path=URLDecoder.decode(path,"utf-8");//转化成utf-8编码
        } catch (UnsupportedEncodingException e) {
        	System.out.println("转化成utf-8编码");
        }
		
		return path;
    }
}
