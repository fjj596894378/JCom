/**
 * 
 */
package cn.com.jcom.start;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import cn.com.jcom.model.WordModel;
import cn.com.jcom.service.ReadExcel;
import cn.com.jcom.service.WriteWord;

/**
 * <P>
 * ��˵����<br/>
 * ˵��д���
 *
 * 
 * </p>
 * �����ߣ�@author ( admin)<br/>
 * ����ʱ�䣺 2016-1-25<br/>
 * 
 */

/**
 * <P>
 * ��˵����<br/>
 * ˵��д���
 * 
 * 
 * </p>
 * �����ߣ�@author ( admin)<br/>
 * ����ʱ�䣺 2016-1-25<br/>
 */
public class RunMain{
	public static void main(String[] args) {

		//��ӳ����������  
		String JarPath = System.getProperty("user.dir");//Util.getCurrentJarPath(); // ��õ�ǰ�����jar��·��
		// xls
		// xlsx

		// File file = new File(str);
		File file = new File(JarPath + "\\1.xlsx");

		if (!file.exists()) {
			// ���������
			file = new File(JarPath + "\\1.xls");
			if (!file.exists()) {
				System.out.println("�ļ�������,�˳�");
				System.exit(0); // ϵͳ�˳�
			}
		}  

		ReadExcel re = new ReadExcel();
		try {
			Thread.sleep(1500);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
		try {

			List<WordModel> list = re.readExcel(file.getAbsolutePath()); // �������ļ�·��
			System.out.println(list);
			int i = 1;
			boolean isDocx = true;
			String filepathString = JarPath + "\\2.docx";
			File fileOut = new File(filepathString);
			if (!fileOut.exists()) {
				isDocx = false;
				fileOut = new File(JarPath + "\\2.doc");
				if (!file.exists()) {
					System.out.println("ģ�治����,�˳�");
					System.exit(0); // ϵͳ�˳�
				}
			} 
			 String destpathString = JarPath + "\\build\\" + i + "ttt.docx";
			  Map<String, String> map = new HashMap<String, String>();
			      String DATE_FORMAT = "yyyy-MM-dd";
			      SimpleDateFormat sf = new SimpleDateFormat(DATE_FORMAT);
			for(WordModel wm : list){
				destpathString = getReturnPath(JarPath,i);
				++i;
				System.out.println(wm.getNo());
				System.out.println(wm.getResult());
				System.out.println(wm.getSubmitTime());
				System.out.println(wm.getSubmitMan());
				System.out.println(wm.getSubmitQuestion());
				System.out.println("--------------------------");
				if(wm.getNo() == null || "".equals(wm.getNo())){
					return ;
				}
				map.put("${NO}", Math.round(Float.valueOf(wm.getNo())) + "");
		        map.put("${submitMan}", wm.getSubmitMan());
		       
		        map.put("${submitTime}", wm.getSubmitTime());
		        map.put("${submitQuestion}", "��������:" +wm.getSubmitQuestion());
		        map.put("${xpSay}", "�����Ӱ�������" + wm.getXpSay());
		        
		        map.put("${level}", wm.getLevel());
		        System.out.println(WriteWord.replaceAndGenerateWord(filepathString,
		                destpathString, map));
			}
			
			
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e2) {
			e2.printStackTrace();
		}
	}

	public static String getReturnPath(String JarPath,int i){
		 String destpathString = JarPath + "\\build\\" + i + "ttt.docx";
		 return destpathString;
	}
	
}
