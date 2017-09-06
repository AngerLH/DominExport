import lotus.domino.*;

import java.io.*;
import java.util.Iterator;
import java.util.Vector;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class JavaAgent extends AgentBase {
	public PublicFunction F = new PublicFunction();
	private Document doc = null;
	private Database db = null;

	public void NotesMain() {
		System.out.println("列表导出中。。。。。。。。。");
		Session session = null;
		AgentContext agentContext = null;
		Agent agent = null;
		String agentName = "";

		View view = null;
		ViewEntryCollection vc = null;
		ViewEntry ve1 = null;
		ViewEntry ve2 = null;
		Document tdoc = null;
		PrintWriter pw = null;

		try {
			session = getSession();
			agentContext = session.getAgentContext();
			db = agentContext.getCurrentDatabase();
			doc = agentContext.getDocumentContext();
			agent = agentContext.getCurrentAgent();
			agentName = agent.getName();
			System.out.println(agentName);
			pw = getAgentOutput();
			vc = getvc();
			System.out.println("vc:" + vc.getCount());
			if (vc != null) {
				String tmpDir = F.GetSysdictionary(db, "ExcelDownload");
				System.out.println(tmpDir);
				XSSFWorkbook workBook = new XSSFWorkbook();
				XSSFSheet sheet = workBook.createSheet();// 创建一个工作薄对象
				XSSFRow row = sheet.createRow(0);// 创建一个行对象
				XSSFCell cell = row.createCell(0);// 创建单元格
				cell.setCellValue("序号");
				XSSFCell cell2 = row.createCell(1);// 创建单元格
				cell2.setCellValue("业务员编号");
				row.createCell(2).setCellValue("业务员名称");
				row.createCell(3).setCellValue("客户名称");
				row.createCell(4).setCellValue("身份证");
				row.createCell(5).setCellValue("电话");
				row.createCell(6).setCellValue("车辆类型");
				row.createCell(7).setCellValue("车型");
				row.createCell(8).setCellValue("车价");
				int i = 1;
				ve1 = vc.getFirstEntry();
				while (ve1 != null) {
					tdoc = ve1.getDocument();
					// 读取每个编制文档的成本中心和科目出来，放到excel的行和单元格里
					XSSFRow row2 = sheet.createRow(i);
					row2.createCell(0).setCellValue(i);
					XSSFCell cell5 = row2.createCell(1);
					cell5.setCellValue(i);

					XSSFCell cell6 = row2.createCell(2);
					cell6.setCellValue(tdoc.getItemValueString("REQUESTNUMBER"));

					XSSFCell cell7 = row2.createCell(3);
					cell7.setCellValue(tdoc.getItemValueString("Salesman_Show"));

					XSSFCell cell8 = row2.createCell(4);
					cell8.setCellValue(tdoc.getItemValueString("PeopleName"));

					XSSFCell cell9 = row2.createCell(5);
					cell9.setCellValue(tdoc.getItemValueString("IDcard"));

					XSSFCell cell10 = row2.createCell(6);
					cell10.setCellValue(tdoc.getItemValueString("PhoneNumber"));

					XSSFCell cell11 = row2.createCell(7);
					cell11.setCellValue(tdoc.getItemValueString("CarType"));

					XSSFCell cell12 = row2.createCell(8);
					cell12.setCellValue(tdoc.getItemValueString("Model"));

					XSSFCell cell13 = row2.createCell(9);
					cell13.setCellValue(tdoc.getItemValueString("CarPrice"));

					i++;
					tdoc.recycle();
					ve2 = vc.getNextEntry(ve1);
					ve1.recycle();
					ve1 = ve2;
				}
				// 循环结束，保存excel文件，并跳转到下载
				System.out.println("新建文件");

				FileOutputStream os = new FileOutputStream(tmpDir
						+ "列表.xlsx");
				workBook.write(os);// 将文档对象写入文件输出流
				os.close();// 关闭文件输出流
				tmpDir = tmpDir.split("html")[1].replace("\\", "/");
				pw.println("<script type='text/javascript'>window.open('"
						+ tmpDir + "列表.xlsx','_blank');</script>");
				pw
						.println("<script type=\"text/javascript\">history.go(-1);</script>");
			} else {
				// 如果没有数据，则alert提示
				pw
						.println("<script type=\"text/javascript\">alert(\"系统不存在编制信息。\");</script>");
			}

		} catch (Exception e) {
			e.printStackTrace();
			String errMsg = getExceptionMsg(e);
			F.printErrMsg(e, doc, errMsg);
		} finally {
			try {
				if (tdoc != null) {
					tdoc.recycle();
				}
				if (ve1 != null) {
					ve1.recycle();
				}
				if (ve2 != null) {
					ve2.recycle();
				}
				if (vc != null) {
					vc.recycle();
				}
				if (view != null) {
					view.recycle();
				}

				if (doc != null) {
					doc.recycle();
				}
				if (db != null) {
					db.recycle();
				}
				if (agent != null) {
					agent.recycle();
				}
				if (agentContext != null) {
					agentContext.recycle();
				}
				if (session != null) {
					session.recycle();
				}
			} catch (NotesException ne) {
				ne.printStackTrace();
			}
		}
	}

	private String getExceptionMsg(Exception e) {
		StackTraceElement[] stes = e.getStackTrace();
		String ret = "";
		for (int i = 0; i < stes.length; i++) {
			ret += stes[i].toString() + "<br>";
		}
		return ret;
	}

	private ViewEntryCollection getvc() {
		try {
			View view = db.getView("v_key_draft2");
			ViewEntryCollection vc = null;
			String budget_id = doc.getItemValueString("budget_id");
			String saleName = doc.getItemValueString("SaleName");
			String date = "";
			String Formula = " FIELD Form CONTAINS " + "mainform"
					+ " and NOT FIELD sys_SoftDelete CONTAINS 1";
			if (Integer.parseInt(doc.getItemValueString("ssMonth")) >= 10) {
				date = doc.getItemValueString("ssYear") + "-"
						+ doc.getItemValueString("ssMonth");
			} else {
				date = doc.getItemValueString("ssYear") + "-0"
						+ doc.getItemValueString("ssMonth");
			}
			vc = view.getAllEntries();

			if (budget_id != "" || saleName != "") {
				Formula = Formula + " and FIELD TDRQ CONTAINS " + date;
				if (budget_id != "") {
					Formula = Formula + " and FIELD SREQUESTNUMBER CONTAINS "
							+ budget_id;
				}
				if (saleName != "") {
					Formula = Formula + " and FIELD Salesman_Show CONTAINS "
							+ saleName;
				}
				if (vc.getCount() > 0) {
					vc.FTSearch(Formula, 0);
				}
			}

			return vc;
		} catch (Exception e2) {
			e2.printStackTrace();
			return null;
		}

	}

}
