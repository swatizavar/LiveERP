package driverFactory;

import commonFunctions.FunctionLibrary;
import utilities.ExcelFileUtil;

public class DriverScript extends FunctionLibrary {
String inputpath="./FileInput/DataEngine.xlsx";
String outputpath="./FileOutput/HybridResults.xlsx";

public void startTest() throws Throwable
{
	String Module_Status="";
	//call excel file util class methods
	ExcelFileUtil xl= new ExcelFileUtil(inputpath);
	//iterate all rows in mastertestcase sheet
	for (int i=1;i<=xl.rowCount("MasterTestCases");i++)
	{
		if(xl.getCellData("MasterTestCases", i, 2).equalsIgnoreCase("Y"))
		{
			//store corresponding sheet names into variable
			String TCModule= xl.getCellData("MasterTestCases", i, 1);
			//iterate all rows in TCModule sheet
			for(int j=1;j<=xl.rowCount(TCModule);j++)
			{
				//call all cellsfrom tcmodule
				String Description=xl.getCellData(TCModule, j, 0);
				String ObjectType=xl.getCellData(TCModule, j, 1);
				String LocatorType=xl.getCellData(TCModule, j, 2);
				String LocatorValue=xl.getCellData(TCModule, j, 3);
				String TestData=xl.getCellData(TCModule, j, 4);
				try {
					if(ObjectType.equalsIgnoreCase("startBrowser"))
					{
						driver=FunctionLibrary.startBrowser();
					}
					else if(ObjectType.equalsIgnoreCase("openUrl"))
					{
						FunctionLibrary.openUrl(driver);
					}
					else if(ObjectType.equalsIgnoreCase("waitForElement"))
					{
						FunctionLibrary.waitForElement(driver, LocatorType, LocatorValue, TestData);
					}
					else if(ObjectType.equalsIgnoreCase("typeAction"))
					{
						FunctionLibrary.typeAction(driver, LocatorType, LocatorValue, TestData);
					}
					else if(ObjectType.equalsIgnoreCase("clickAction"))
					{
						FunctionLibrary.clickAction(driver, LocatorType, LocatorValue);
					}
					else if(ObjectType.equalsIgnoreCase("validateTitle"))
					{
						FunctionLibrary.validateTitle(driver,TestData);
					}
					else if(ObjectType.equalsIgnoreCase("closeBrowser"))
					{
						FunctionLibrary.closeBrowser(driver);
					}
					//write as pass into status cell TCModule 
					xl.setCellData(ObjectType, j, 5, "Pass", outputpath);
					Module_Status="True";
				}
				catch(Exception e) {
					System.out.println(e.getMessage());
					//write as fail into status cell TCModule
					xl.setCellData(ObjectType, j, 5, "Fail", outputpath);
					Module_Status="False";

				}
				// if all the steps pass then we write pass in test cases sheet
				if(Module_Status.equalsIgnoreCase("True"))
				{
					xl.setCellData("MasterTestCases", i, 3, "Pass", outputpath);
				}
				// If not write as fail
				else
				{
					xl.setCellData("MasterTestCases", i, 3, "Fail", outputpath);
				}
			}
		}
		else
		{
			//Write as Blocked in status column which are flaged to N
			xl.setCellData("MasterTestCase", i, 3, "Blocked", outputpath);
		}
	}
}
}
