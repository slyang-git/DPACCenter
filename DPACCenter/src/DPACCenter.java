import java.io.File;



public class DPACCenter {

	public static void main(String[] args) throws Exception {
		XSLXFileFilter filter = new XSLXFileFilter();
		File dir = new File(".");
		File[] filename = dir.listFiles(filter);
		//逐个文件处理
		for (File f : filename) {
			String name = f.getName();
			//System.out.println(name);
			if (name.contains("美国")){
				American american = new American();
				american.XiaoFeiPinRecall(name);
				american.QiCheRecall(name);
				
			}else if(name.contains("欧盟")){	//处理欧盟召回信息表
				EuropeanUnion europeanUnion = new EuropeanUnion();
				europeanUnion.XiaoFeiPinRecall(name);
				europeanUnion.QiCheRecall(name);
				
			}else if(name.contains("澳大利亚")){
				Australia australia = new Australia();
				australia.XiaoFeiPinRecall(name);
				australia.QiCheRecall(name);
				
			}else if(name.contains("韩国")){
				Korea korea = new Korea();
				korea.QiCheRecall(name);
				
			}else if(name.contains("英国")){
				Britain britan = new Britain();
				britan.QiCheRecall(name);
				
			}else if(name.contains("日本")){
				Japan japan = new Japan();
				if (name.contains("汽车")){
					japan.QiCheRecall(name);
				}else if (name.contains("消费品")){
					japan.XiaoFeiPinRecall(name);
				}
				
			}else if(name.contains("德国")){
				Germany germany = new Germany();
				germany.QiCheRecall(name);
			}
		}
		System.out.println("<---------程序运行完成了------------>");
	}//for (File f : filename)

}
