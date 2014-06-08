import java.io.File;
import java.io.FilenameFilter;


public class XSLXFileFilter implements FilenameFilter {
	
	String fileType = ".xlsx"; //�ļ�����
	@Override
	public boolean accept(File dir, String name) {
		return (name.endsWith(fileType) && !name.startsWith("~$"));
	}

}
