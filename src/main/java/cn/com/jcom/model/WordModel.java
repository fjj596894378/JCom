/**
 * 
 */
package cn.com.jcom.model;

import java.io.Serializable;


/**
 * <P>
 * ��˵����<br/>
 * ˵��д���
 *
 * 
 * </p>
 * �����ߣ�@author ( admin)<br/>
 * ����ʱ�䣺 2016-1-27<br/>
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
 * ����ʱ�䣺 2016-1-27<br/>
 */
public class WordModel implements Serializable{
	/**
	 * �ֶ�˵��:xxxx <br/>
	 */
	private static final long serialVersionUID = 1L;
	private String no;                 // ���
	private String submitTime;	// ���ʱ��
	private String submitMan;	// �ĸ���B�������
	private String submitQuestion;	// �ĸ���B����
	private String submitType;	// ��������
	private String isback;	// �Ƿ���
	private String isDealWith;	// �Ƿ��Ѿ�����
	private String result;	// �������
	private String xpSay; // л���ش�
	private String level; // �ȼ�
	/**
	 * @return the no
	 */
	public String getNo() {
		return no;
	}
	/**
	 * @param no the no to set
	 */
	public void setNo(String no) {
		this.no = no;
	}
	/**
	 * @return the submitTime
	 */
	public String getSubmitTime() {
		return submitTime;
	}
	/**
	 * @param submitTime the submitTime to set
	 */
	public void setSubmitTime(String submitTime) {
		this.submitTime = submitTime;
	}
	/**
	 * @return the submitMan
	 */
	public String getSubmitMan() {
		return submitMan;
	}
	/**
	 * @param submitMan the submitMan to set
	 */
	public void setSubmitMan(String submitMan) {
		this.submitMan = submitMan;
	}
	/**
	 * @return the submitQuestion
	 */
	public String getSubmitQuestion() {
		return submitQuestion;
	}
	/**
	 * @param submitQuestion the submitQuestion to set
	 */
	public void setSubmitQuestion(String submitQuestion) {
		this.submitQuestion = submitQuestion;
	}
	/**
	 * @return the submitType
	 */
	public String getSubmitType() {
		return submitType;
	}
	/**
	 * @param submitType the submitType to set
	 */
	public void setSubmitType(String submitType) {
		this.submitType = submitType;
	}
	/**
	 * @return the isback
	 */
	public String getIsback() {
		return isback;
	}
	/**
	 * @param isback the isback to set
	 */
	public void setIsback(String isback) {
		this.isback = isback;
	}
	/**
	 * @return the isDealWith
	 */
	public String getIsDealWith() {
		return isDealWith;
	}
	/**
	 * @param isDealWith the isDealWith to set
	 */
	public void setIsDealWith(String isDealWith) {
		this.isDealWith = isDealWith;
	}
	/**
	 * @return the result
	 */
	public String getResult() {
		return result;
	}
	/**
	 * @param result the result to set
	 */
	public void setResult(String result) {
		this.result = result;
	}
	/**
	 * @return the xpSay
	 */
	public String getXpSay() {
		return xpSay;
	}
	/**
	 * @param xpSay the xpSay to set
	 */
	public void setXpSay(String xpSay) {
		this.xpSay = xpSay;
	}
	/**
	 * @return the level
	 */
	public String getLevel() {
		return level;
	}
	/**
	 * @param level the level to set
	 */
	public void setLevel(String level) {
		this.level = level;
	}
	
	
}