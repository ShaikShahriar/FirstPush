package utilities;

import javax.imageio.stream.FileImageInputStream;
import java.util.Properties;

public class PropertFileUtils{
        public static String getValueForKey(String key)throws Throwable
        {
            Properties conpro = new Properties();
            //load file
            conpro.load(new FileImageInputStream("./LiveProject\\Environment.properties"));
            return conpro.getProperty(key);
            

        }

}
