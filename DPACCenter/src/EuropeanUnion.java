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

//�W��
public class EuropeanUnion {
	
	private String biaoti; //����
	private String faburiqi; //��������
	private String fabuguojia = "�ٻط������һ����:ŷ��"; //��������
	private String chanpinmingchen; //��Ʒ����
	private String chandi; //����
	private String jutixinghao; //�����ͺ�
	private String quexianhouguo; //ȱ�ݼ����
	private String zhizaoshang;	//������
	private String shuliang; //�ٻ�����
	private String anquantixing = "�����ʼ��ܾ�ȱ�ݲ�Ʒ����������ʾ��������������Ĳ�Ʒ�����������⣬���Է��ʱ���վ��ȱ�ݲɼ�����Ŀ��http://www.dpac.gov.cn���ύ��ϸ��Ϣ�����߲���010-59799616������ѯ��"; //��ȫ����
	
	//����ŷ�������ٻ���Ϣ���ݱ�
	public void QiCheRecall(String filename) throws Exception {
		//������
		FileInputStream in = new FileInputStream(filename);
		XSSFWorkbook wb = new XSSFWorkbook(in);
		//�����
		OutputStreamWriter fow = new OutputStreamWriter(new FileOutputStream(filename.replaceAll(".xlsx", "") + "��������" + ".txt"),"GBK");
		BufferedWriter writer = new BufferedWriter(fow);
		// ��õ�һ�ű�
		XSSFSheet sheet = wb.getSheetAt(1);
		for (int i = 2; i <= sheet.getLastRowNum(); i++) {
			XSSFRow row = sheet.getRow(i); // ��õ�i��
			if (row.getCell(2) != null) {
				XSSFCellStyle cellstyle = row.getRowStyle();
				if (cellstyle != null) {
					XSSFColor color = cellstyle.getFillBackgroundColorColor();
					if (color != null) {
						if (color.getTheme() == 0) { //���б������ɫ����
							XSSFCell cell = row.getCell(3); //����
							chanpinmingchen = "��Ʒ����: " + cell.getStringCellValue();
							biaoti = "��ŷ�ˡ�" + cell.getStringCellValue().trim()+"�����ٻ�";
							cell = row.getCell(0);
							SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
							faburiqi = "�������ڣ�" + sdf.format(cell.getDateCellValue());
							cell = row.getCell(2);
							zhizaoshang = "�����̣�" + cell.getStringCellValue();
							cell = row.getCell(4);
							jutixinghao = "���ͣ�" + cell.getStringCellValue();
							if ( (cell = row.getCell(5)) !=null){
								shuliang = cell.getStringCellValue();
							}else{
								shuliang = "����������";
							}
							cell = row.getCell(8);
							quexianhouguo = "ȱ�ݼ����: " + cell.getStringCellValue();
							
							
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
	
	//����ŷ��һ������Ʒ�ٻ���Ϣ���ݱ�
	public void XiaoFeiPinRecall(String filename) throws Exception {
				//������
				FileInputStream in = new FileInputStream(filename);
				XSSFWorkbook wb = new XSSFWorkbook(in);
				//�����
				OutputStreamWriter fow = new OutputStreamWriter(new FileOutputStream(filename.replaceAll(".xlsx", "") + "��һ������Ʒ��" + ".txt"),"GBK");
				BufferedWriter writer = new BufferedWriter(fow);
				// ��õ�һ�ű�
				XSSFSheet sheet = wb.getSheetAt(0);
				for (int i = 2; i <= sheet.getLastRowNum(); i++) {
					XSSFRow row = sheet.getRow(i); // ��õ�i��
					if (row.getCell(2) != null) {
						XSSFCellStyle cellstyle = row.getRowStyle();
						if (cellstyle != null) {
							XSSFColor color = cellstyle.getFillBackgroundColorColor();
							if (color != null) {
								if (color.getTheme() == 0) { //���б������ɫ����
									XSSFCell cell = row.getCell(3); //����
									chanpinmingchen = "��Ʒ����: " + cell.getStringCellValue();
									biaoti = "��ŷ�ˡ�" + cell.getStringCellValue().trim()+"�ٻ�";
									cell = row.getCell(0);
									SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
									faburiqi = "�ٻط������ڣ�" + sdf.format(cell.getDateCellValue());
									cell = row.getCell(6);
									jutixinghao = "�����ͺŻ�ʶ��������" + cell.getStringCellValue();
									cell = row.getCell(12);
									quexianhouguo = "ȱ�ݼ����: " + cell.getStringCellValue();
									
									cell = row.getCell(7);
									if (cell != null && !cell.getStringCellValue().isEmpty()) {
										chandi = "���أ�" + cell.getStringCellValue();
									} else if((cell = row.getCell(8) ) != null && !cell.getStringCellValue().isEmpty()){
										chandi = "���أ�" + cell.getStringCellValue();
									}else{
										chandi = "���أ�����";
									}
									writer.write(biaoti + "\n\r");
									writer.write(faburiqi + "\n\r");
									writer.write(fabuguojia + "\n\r");
									writer.write(chanpinmingchen + "\n\r");
									writer.write(chandi + "\n\r");
									writer.write(jutixinghao + "\n\r");
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
