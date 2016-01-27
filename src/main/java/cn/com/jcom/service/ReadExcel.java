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
 * 类说明：<br/>
 * 说明写这儿
 *
 * 
 * </p>
 * 创建者：@author ( fujj)<br/>
 * 创建时间： 2016-1-27<br/>
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
                    XSSFCell no = xssfRow.getCell(0); // 序号
                    XSSFCell submitTime = xssfRow.getCell(1); // 提出时间
                    XSSFCell submitMan = xssfRow.getCell(2); // 哪个逗B提的问题
                    XSSFCell submitQuestion = xssfRow.getCell(3); // 哪个逗B问题
                    XSSFCell submitType = xssfRow.getCell(4);	// 问题类型
                    XSSFCell isback = xssfRow.getCell(5);	// 是否反馈
                    XSSFCell isDealWith = xssfRow.getCell(6);	// 是否已经处理
                    XSSFCell result = xssfRow.getCell(7);	 // 反馈结果
                    XSSFCell xpSay = xssfRow.getCell(8);	 // 谢鹏说
                    XSSFCell level = xssfRow.getCell(9);	 // 谢鹏说
                    wm.setNo(getValue(no));
                    wm.setSubmitTime(parseExcel(submitTime));
                    wm.setSubmitMan(getValue(submitMan));
                    wm.setSubmitQuestion(getValue(submitQuestion));
                    wm.setSubmitType(getValue(submitType));
                    wm.setIsback(getValue(isback));
                    wm.setIsDealWith(getValue(isDealWith));
                    wm.setResult(getValue(result));
                    wm.setXpSay(getValue(xpSay));
                    if(getValue(level).equals("高")){
                    	wm.setLevel("口低 口中√高");
                    }else if(getValue(level).equals("低")){
                    	wm.setLevel("√低 口中口高");
                    }else{
                    	wm.setLevel("口低 √中口高");
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
                    HSSFCell no = hssfRow.getCell(0); // 序号
                    HSSFCell submitTime = hssfRow.getCell(1); // 提出时间
                   
                    HSSFCell submitMan = hssfRow.getCell(2); // 哪个逗B提的问题
                    HSSFCell submitQuestion = hssfRow.getCell(3); // 哪个逗B问题
                    HSSFCell submitType = hssfRow.getCell(4);	// 问题类型
                    HSSFCell isback = hssfRow.getCell(5);	// 是否反馈
                    HSSFCell isDealWith = hssfRow.getCell(6);	// 是否已经处理
                    HSSFCell result = hssfRow.getCell(7);	 // 反馈结果
                    HSSFCell xpSay = hssfRow.getCell(8);	 // 谢鹏说
                    HSSFCell level = hssfRow.getCell(9);	 // 谢鹏说
                    wm.setNo(getValue(no));
                    wm.setSubmitTime( parseExcel(submitTime));
                    wm.setSubmitMan(getValue(submitMan));
                    wm.setSubmitQuestion(getValue(submitQuestion));
                    wm.setSubmitType(getValue(submitType));
                    wm.setIsback(getValue(isback));
                    wm.setIsDealWith(getValue(isDealWith));
                    wm.setResult(getValue(result));
                    wm.setXpSay(getValue(xpSay));
                    if(getValue(level).equals("高")){
                    	wm.setLevel("口低 口中√高");
                    }else if(getValue(level).equals("低")){
                    	wm.setLevel("√低 口中口高");
                    }else{
                    	wm.setLevel("口低 √中口高");
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
		case HSSFCell.CELL_TYPE_NUMERIC:// 数字类型
			if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
				SimpleDateFormat sdf = null;
				if (cell.getCellStyle().getDataFormat() == HSSFDataFormat
						.getBuiltinFormat("h:mm")) {
					sdf = new SimpleDateFormat("HH:mm");
				} else {// 日期
					sdf = new SimpleDateFormat("yyyy-MM-dd");
				}
				Date date = cell.getDateCellValue();
				result = sdf.format(date);
			} else if (cell.getCellStyle().getDataFormat() == 58) {
				// 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
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
				// 单元格设置成常规
				if (temp.equals("General")) {
					format.applyPattern("#");
				}
				result = format.format(value);
			}
			break;
		case HSSFCell.CELL_TYPE_STRING:// String类型
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