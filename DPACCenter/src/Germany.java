import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.text.SimpleDateFormat;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Germany {
	private String biaoti; // 标题
	private String faburiqi; // 发布日期
	private String fabuguojia = "召回发布国家或地区:德国"; // 发布国家
	private String chanpinmingchen; // 产品名称
	private String jutixinghao; // 具体型号
	private String quexianhouguo; // 缺陷及后果
	private String zhizaoshang; // 制造商
	private String shuliang; // 召回数量
	private String anquantixing = "国家质检总局缺陷产品管理中心提示：如果您发现您的产品出现类似问题，可以访问本网站“缺陷采集”栏目（http://www.dpac.gov.cn）提交详细信息，或者拨打010-59799616进行咨询。"; // 安全提醒

	// 处理汽车召回表
	public void QiCheRecall(String filename) throws Exception {
		//输入流
		FileInputStream in = new FileInputStream(filename);
		XSSFWorkbook wb = new XSSFWorkbook(in);
		//输出流
		OutputStreamWriter fow = new OutputStreamWriter(new FileOutputStream(filename.replaceAll(".xlsx", "") + "（汽车）" + ".txt"),"GBK");
		BufferedWriter writer = new BufferedWriter(fow);
		// 获得第一张表
		XSSFSheet sheet = wb.getSheetAt(0);
		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			XSSFRow row = sheet.getRow(i); // 获得第i行
			if (row.getCell(2) != null) {
				XSSFCellStyle cellstyle = row.getRowStyle();
				if (cellstyle != null) {
					XSSFColor color = cellstyle.getFillBackgroundColorColor();
					if (color != null) {
						if (color.getTheme() == 0) { //具有背景填充色的行
							XSSFCell cell = row.getCell(3); //标题
							if (cell != null){
								chanpinmingchen = cell.getStringCellValue();
							}else {
								cell = row.getCell(2);
								chanpinmingchen = cell.getStringCellValue();
							}
							
							biaoti = "【德国】" + chanpinmingchen.trim()+"召回";
							cell = row.getCell(0);
							SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
							faburiqi = "发布日期：" + sdf.format(cell.getDateCellValue());
							//faburiqi = "发布日期：" + cell.getStringCellValue();
							cell = row.getCell(2);
							zhizaoshang = "制造商：" + cell.getStringCellValue();
							cell = row.getCell(4);
							if(cell != null){
								jutixinghao = "车型：" + cell.getStringCellValue();
							}else {
								jutixinghao = "车型：不详";
							}
							
							if ( (cell = row.getCell(5)) !=null){
								switch(cell.getCellType()){
								case XSSFCell.CELL_TYPE_NUMERIC:
									shuliang = "召回数量：" + (int)cell.getNumericCellValue();
									break;
								case XSSFCell.CELL_TYPE_STRING:
									shuliang = "召回数量：" + cell.getStringCellValue();
									break;
								}
								//shuliang = "召回数量：" + cell.getStringCellValue();
							}else{
								shuliang = "召回数量：不详";
							}
							cell = row.getCell(8);
							quexianhouguo = "缺陷及后果: " + cell.getStringCellValue();
							
							
							writer.write(biaoti + "\n\r");
							writer.write(faburiqi + "\n\r");
							writer.write(fabuguojia + "\n\r");
							writer.write(zhizaoshang + "\n\r");
							writer.write(jutixinghao + "\n\r");
							writer.write(shuliang + "\n\r");
							writer.write(quexianhouguo + "\n\r");
							writer.write(anquantixing + "\n\r");
							
							writer.newLine();
						}// if ( color.getTheme() == 0)
					}// if (color != null)
				}
			}

		}// for (int i=0; i<sheet.getLastRowNum(); i++)

		in.close();
		writer.close();
		fow.close();

	}
}
