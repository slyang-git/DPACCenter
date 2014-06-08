import java.io.File;



public class DPACCenter {

	public static void main(String[] args) throws Exception {
		XSLXFileFilter filter = new XSLXFileFilter();
		File dir = new File(".");
		File[] filename = dir.listFiles(filter);
		//����ļ�����
		for (File f : filename) {
			String name = f.getName();
			//System.out.println(name);
			if (name.contains("����")){
				American american = new American();
				american.XiaoFeiPinRecall(name);
				american.QiCheRecall(name);
				
			}else if(name.contains("ŷ��")){	//����ŷ���ٻ���Ϣ��
				EuropeanUnion europeanUnion = new EuropeanUnion();
				europeanUnion.XiaoFeiPinRecall(name);
				europeanUnion.QiCheRecall(name);
				
			}else if(name.contains("�Ĵ�����")){
				Australia australia = new Australia();
				australia.XiaoFeiPinRecall(name);
				australia.QiCheRecall(name);
				
			}else if(name.contains("����")){
				Korea korea = new Korea();
				korea.QiCheRecall(name);
				
			}else if(name.contains("Ӣ��")){
				Britain britan = new Britain();
				britan.QiCheRecall(name);
				
			}else if(name.contains("�ձ�")){
				Japan japan = new Japan();
				if (name.contains("����")){
					japan.QiCheRecall(name);
				}else if (name.contains("����Ʒ")){
					japan.XiaoFeiPinRecall(name);
				}
				
			}else if(name.contains("�¹�")){
				Germany germany = new Germany();
				germany.QiCheRecall(name);
			}
		}
		System.out.println("<---------�������������------------>");
	}//for (File f : filename)

}
