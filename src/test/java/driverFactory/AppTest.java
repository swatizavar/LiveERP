package driverFactory;

import org.testng.annotations.Test;

public class AppTest {
//This class is created to execute as maven project
	@Test
	public void kickStart() throws Throwable
	{
		DriverScript ds=new DriverScript();
		ds.startTest();
	}
	// next right click on pom.xml run as maven test to run project
}
