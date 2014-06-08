import java.io.File;
import java.io.FilenameFilter;


public class XSLXFileFilter implements FilenameFilter {
	
	String fileType = ".xlsx"; //文件类型
	@Override
	public boolean accept(File dir, String name) {
		return (name.endsWith(fileType) && !name.startsWith("~$"));
	}

}
