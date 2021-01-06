package com.setUp;

import java.io.FileReader;
import java.util.Properties;

public   class setUp 
{
	public static Properties envSetUp() 
	{
		Properties prop = new Properties();
		try {
			String envPath = System.getProperty("user.dir") + "\\src\\main\\resources\\config.properties";

			FileReader reader = new FileReader(envPath);
			prop.load(reader);

		} catch (Exception e) {
			System.out.println(e.getMessage());
		}

		return prop;
	}

}
