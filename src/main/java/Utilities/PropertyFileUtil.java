package Utilities;

import java.io.FileInputStream;
import java.util.Properties;

public class PropertyFileUtil {
	public static String getValueForKey(String key)throws Throwable {
		Properties confingProperties =new Properties();
		confingProperties.load(new FileInputStream("D:\\practice\\ERP_Maven\\PropertyFile\\Environment.properties"));
		return confingProperties.getProperty(key);
	}
}
