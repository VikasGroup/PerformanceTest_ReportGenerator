package com.Excel;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;



public class PropertyReader {
	public Properties propp = new Properties();

	public PropertyReader() {
		try {
			InputStream file = getClass().getResourceAsStream("/testdata/data.properties");
			propp.load(file);

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
