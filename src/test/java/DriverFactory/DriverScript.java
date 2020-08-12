package DriverFactory;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import CommonFunLibrary.FunctionLibary;
import Utilities.ExcelfileUtil;

public class DriverScript {
	String inputpath="D:\\practice\\ERP_Maven\\TestInput\\InputSheet.xlsx";
	String outputpath="D:\\practice\\ERP_Maven\\TestOutput\\hybrid.xlsx";
	WebDriver driver;
	ExtentReports report;
	ExtentTest test;
	public void startTest()throws Throwable {
		//creating referance object for excel util method
		ExcelfileUtil xl=new ExcelfileUtil(inputpath);
		//itterating all row in mastertestcase sheet
		for(int i=1;i<=xl.rowCount("MasterTestCases");i++) {
			if(xl.getCellData("MasterTestCases", i, 2).equalsIgnoreCase("Y")) {
				//store module name into TCModule
				String TCModule=xl.getCellData("MasterTestCases", i, 1);
				report =new ExtentReports("./ExtentReports/"+TCModule+FunctionLibary.generateDate()+".html");
				//ittirate all row in TCModule
				for(int j=1;j<=xl.rowCount(TCModule);j++) {
					//test case start here
					test=report.startTest(TCModule);
					//read cell from TCModule
					String Description=xl.getCellData(TCModule, j, 0);
					String Function_Name=xl.getCellData(TCModule, j, 1);
					String Locator_Type=xl.getCellData(TCModule, j, 2);
					String Locator_Value=xl.getCellData(TCModule, j, 3);
					String Test_Data=xl.getCellData(TCModule, j, 4);
					try {
						//call function
						if(Function_Name.equalsIgnoreCase("startBrowser")) {
							driver=FunctionLibary.startBrowser();
							test.log(LogStatus.INFO, Description);
						}
						else if(Function_Name.equalsIgnoreCase("openApplication")) {
							FunctionLibary.openApplication(driver);
							test.log(LogStatus.INFO, Description);
						}
						else if(Function_Name.equalsIgnoreCase("waitForElement")) {
							FunctionLibary.waitForElement(driver, Locator_Type, Locator_Value, Test_Data);
							test.log(LogStatus.INFO, Description);
						}
						else if(Function_Name.equalsIgnoreCase("typeAction")) {
							FunctionLibary.typeAction(driver, Locator_Type, Locator_Value, Test_Data);
							test.log(LogStatus.INFO, Description);
						}
						else if(Function_Name.equalsIgnoreCase("clickAction")) {
							FunctionLibary.clickAction(driver, Locator_Type, Locator_Value);
							test.log(LogStatus.INFO, Description);
						}
						else if(Function_Name.equalsIgnoreCase("closeBrowser")) {
							FunctionLibary.closeBrowser(driver);
							test.log(LogStatus.INFO, Description);
						}
						else if(Function_Name.equalsIgnoreCase("captureData")) {
							FunctionLibary.captureData(driver, Locator_Type, Locator_Value);
							test.log(LogStatus.INFO, Description);
						}
						else if(Function_Name.equalsIgnoreCase("tableValidation")) {
							FunctionLibary.tableValidation(driver, Test_Data);
							test.log(LogStatus.INFO, Description);
						}
						else if(Function_Name.equalsIgnoreCase("stockCategories")) {
							FunctionLibary.stockCategories(driver);
							test.log(LogStatus.INFO, Description);
						}
						else if(Function_Name.equalsIgnoreCase("stockValidation")) {
							FunctionLibary.stockValidation(driver, Test_Data);
							test.log(LogStatus.INFO, Description);
						}
						else if(Function_Name.equalsIgnoreCase("captureData1")) {
							FunctionLibary.captureData1(driver, Locator_Type, Locator_Value);
							System.out.println("Data capture sucessfully");
							test.log(LogStatus.INFO, Description);
						}
						else if(Function_Name.equalsIgnoreCase("tableValidation1")) {
							FunctionLibary.tableValidation1(driver, Test_Data);
							test.log(LogStatus.INFO, Description);
						}
						//write as pass into status colume TCModule
						xl.setCellData(TCModule, j, 5, "Pass", outputpath);
						test.log(LogStatus.PASS, Description);
						//write as pass into mastertestcases sheet
						xl.setCellData("MasterTestCases", i, 3, "Pass", outputpath);
					}catch(Exception e) {
						System.out.println(e.getMessage());
						//write as fail into status column TCModule
						xl.setCellData(TCModule, j, 5, "Fail", outputpath);
						test.log(LogStatus.FAIL, Description);
						File screen=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
						FileUtils.copyFile(screen, new File("./Screens/"+TCModule+FunctionLibary.generateDate()+".png"));
						//Write as fail into master testcases sheet
						xl.setCellData("MasterTestCases", i, 3, "Fail", outputpath);
					}
				}
				report.endTest(test);
				report.flush();
			}
			else {
				//write as not executed into status column in master test case sheet
				xl.setCellData("MasterTestCases", i, 3, "Not Executed", outputpath);
			}
		}
	}
}
