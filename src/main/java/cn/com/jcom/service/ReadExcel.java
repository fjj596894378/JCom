/**
 * 
 */
package cn.com.jcom.service;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cn.com.jcom.model.Common;
import cn.com.jcom.model.WordModel;
import cn.com.jcom.util.Util;

/**
 * <P>
 * ��˵����<br/>
 * ˵��д���
 *
 * 
 * </p>
 * �����ߣ�@author ( fujj)<br/>
 * ����ʱ�䣺 2016-1-27<br/>
 */
public class ReadExcel {
	/**
     * read the Excel file
     * @param path the path of the Excel file
     * @return
     * @throws IOException
     */
    public List<WordModel> readExcel(String path) throws IOException {
        if (path == null || Common.EMPTY.equals(path)) {
            return null;
        } else {
            String postfix = Util.getPostfix(path);
            if (!Common.EMPTY.equals(postfix)) {
                if (Common.OFFICE_EXCEL_2003_POSTFIX.equals(postfix)) {
                	 
                    return readXls(path);
                } else if (Common.OFFICE_EXCEL_2010_POSTFIX.equals(postfix)) {
                    return readXlsx(path);
                }
            } else {
                System.out.println(path + Common.NOT_EXCEL_FILE);
            }
        }
        return null;
    }

    /**
     * Read the Excel 2010
     * @param path the path of the excel file
     * @return
     * @throws IOException
     */
    public List<WordModel> readXlsx(String path) throws IOException {
        System.out.println(Common.PROCESSING + path);
        InputStream is = new FileInputStream(path);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
        WordModel wm = null;
        List<WordModel> list = new ArrayList<WordModel>();
        // Read the Sheet
        for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
            if (xssfSheet == null) {
                continue;
            }
            // Read the Row
            for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
                XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                if (xssfRow != null) {
                    wm = new WordModel();
                    XSSFCell no = xssfRow.getCell(0); // ���
                    XSSFCell submitTime = xssfRow.getCell(1); // ���ʱ��
                    XSSFCell submitMan = xssfRow.getCell(2); // �ĸ���B�������
                    XSSFCell submitQuestion = xssfRow.getCell(3); // �ĸ���B����
                    XSSFCell submitType = xssfRow.getCell(4);	// ��������
                    XSSFCell isback = xssfRow.getCell(5);	// �Ƿ���
                    XSSFCell isDealWith = xssfRow.getCell(6);	// �Ƿ��Ѿ�����
                    XSSFCell result = xssfRow.getCell(7);	 // �������
                    XSSFCell xpSay = xssfRow.getCell(8);	 // л��˵
                    XSSFCell level = xssfRow.getCell(9);	 // л��˵
                    wm.setNo(getValue(no));
                    wm.setSubmitTime(parseExcel(submitTime));
                    wm.setSubmitMan(getValue(submitMan));
                    wm.setSubmitQuestion(getValue(submitQuestion));
                    wm.setSubmitType(getValue(submitType));
                    wm.setIsback(getValue(isback));
                    wm.setIsDealWith(getValue(isDealWith));
                    wm.setResult(getValue(result));
                    wm.setXpSay(getValue(xpSay));
                    if(getValue(level).equals("��")){
                    	wm.setLevel("�ڵ� ���С̸�");
                    }else if(getValue(level).equals("��")){
                    	wm.setLevel("�̵� ���пڸ�");
                    }else{
                    	wm.setLevel("�ڵ� ���пڸ�");
                    }
                    list.add(wm);
                }
            }
        }
        return list;
    }

    /**
     * Read the Excel 2003-2007
     * @param path the path of the Excel
     * @return
     * @throws IOException
     */
    public List<WordModel> readXls(String path) throws IOException {
        System.out.println(Common.PROCESSING + path);
        InputStream is = new FileInputStream(path);
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
        WordModel wm = null;
        List<WordModel> list = new ArrayList<WordModel>();
        // Read the Sheet
        for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
            if (hssfSheet == null) {
                continue;
            }
            // Read the Row
            for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                HSSFRow hssfRow = hssfSheet.getRow(rowNum);
                if (hssfRow != null) {
                   /* student = new WordModel();
                    HSSFCell no = hssfRow.getCell(0);
                    HSSFCell name = hssfRow.getCell(1);
                    HSSFCell age = hssfRow.getCell(2);
                    HSSFCell score = hssfRow.getCell(3);
                    student.setNo(getValue(no));
                    student.setName(getValue(name));
                    student.setAge(getValue(age));
                    student.setScore(Float.valueOf(getValue(score)));*/
                    
                    wm = new WordModel();
                    HSSFCell no = hssfRow.getCell(0); // ���
                    HSSFCell submitTime = hssfRow.getCell(1); // ���ʱ��
                   
                    HSSFCell submitMan = hssfRow.getCell(2); // �ĸ���B�������
                    HSSFCell submitQuestion = hssfRow.getCell(3); // �ĸ���B����
                    HSSFCell submitType = hssfRow.getCell(4);	// ��������
                    HSSFCell isback = hssfRow.getCell(5);	// �Ƿ���
                    HSSFCell isDealWith = hssfRow.getCell(6);	// �Ƿ��Ѿ�����
                    HSSFCell result = hssfRow.getCell(7);	 // �������
                    HSSFCell xpSay = hssfRow.getCell(8);	 // л��˵
                    HSSFCell level = hssfRow.getCell(9);	 // л��˵
                    wm.setNo(getValue(no));
                    wm.setSubmitTime( parseExcel(submitTime));
                    wm.setSubmitMan(getValue(submitMan));
                    wm.setSubmitQuestion(getValue(submitQuestion));
                    wm.setSubmitType(getValue(submitType));
                    wm.setIsback(getValue(isback));
                    wm.setIsDealWith(getValue(isDealWith));
                    wm.setResult(getValue(result));
                    wm.setXpSay(getValue(xpSay));
                    if(getValue(level).equals("��")){
                    	wm.setLevel("�ڵ� ���С̸�");
                    }else if(getValue(level).equals("��")){
                    	wm.setLevel("�̵� ���пڸ�");
                    }else{
                    	wm.setLevel("�ڵ� ���пڸ�");
                    }
                    list.add(wm);
                }
            }
        }
        return list;
    }

    @SuppressWarnings("static-access")
    private String getValue(XSSFCell xssfRow) {
    	if(xssfRow == null) {
    		return "";
    	}
        if (xssfRow.getCellType() == xssfRow.CELL_TYPE_BOOLEAN) {
            return String.valueOf(xssfRow.getBooleanCellValue());
        } else if (xssfRow.getCellType() == xssfRow.CELL_TYPE_NUMERIC) {
            return String.valueOf(xssfRow.getNumericCellValue());
        } else {
            return String.valueOf(xssfRow.getStringCellValue());
        }
    }

    @SuppressWarnings("static-access")
    private String getValue(HSSFCell hssfCell) {
    	if(hssfCell == null) {
    		return "";
    	}
        if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
            return String.valueOf(hssfCell.getBooleanCellValue());
        } else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
            return String.valueOf(hssfCell.getNumericCellValue());
        } else {
            return String.valueOf(hssfCell.getStringCellValue());
        }
    }
    
    private String parseExcel(Cell cell) {
		String result = new String();
		switch (cell.getCellType()) {
		case HSSFCell.CELL_TYPE_NUMERIC:// ��������
			if (HSSFDateUtil.isCellDateFormatted(cell)) {// �������ڸ�ʽ��ʱ���ʽ
				SimpleDateFormat sdf = null;
				if (cell.getCellStyle().getDataFormat() == HSSFDataFormat
						.getBuiltinFormat("h:mm")) {
					sdf = new SimpleDateFormat("HH:mm");
				} else {// ����
					sdf = new SimpleDateFormat("yyyy-MM-dd");
				}
				Date date = cell.getDateCellValue();
				result = sdf.format(date);
			} else if (cell.getCellStyle().getDataFormat() == 58) {
				// �����Զ������ڸ�ʽ��m��d��(ͨ���жϵ�Ԫ��ĸ�ʽid�����id��ֵ��58)
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
				double value = cell.getNumericCellValue();
				Date date = org.apache.poi.ss.usermodel.DateUtil
						.getJavaDate(value);
				result = sdf.format(date);
			} else {
				double value = cell.getNumericCellValue();
				CellStyle style = cell.getCellStyle();
				DecimalFormat format = new DecimalFormat();
				String temp = style.getDataFormatString();
				// ��Ԫ�����óɳ���
				if (temp.equals("General")) {
					format.applyPattern("#");
				}
				result = format.format(value);
			}
			break;
		case HSSFCell.CELL_TYPE_STRING:// String����
			result = cell.getRichStringCellValue().toString();
			break;
		case HSSFCell.CELL_TYPE_BLANK:
			result = "";
		default:
			result = "";
			break;
		}
		return result;
	}

}