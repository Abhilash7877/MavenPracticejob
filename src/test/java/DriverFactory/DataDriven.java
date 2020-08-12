package DriverFactory;

import org.openqa.selenium.WebDriver;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import CommonFunLibrary.ERP_Function;
import Utilities.ExcelfileUtil;

public class DataDriven {
	WebDriver driver;
	String inputpath="D:\\practice\\ERP_StockAccounting\\TestInput\\TestData.xlsx";
	String outpath="D:\\practice\\ERP_StockAccounting\\TestOutput\\DataDriven.xlsx";
	@BeforeTest
	public void setup()throws Throwable {
		String Launch=ERP_Function.verifyUrl("http://projects.qedgetech.com/webapp");
		Reporter.log(Launch,true);
		String loginresults=ERP_Function.verifyLogin("admin", "master");
		Reporter.log(loginresults,true);
	}
	@Test
	public void supplier()throws Throwable {
		//access excel method
		ExcelfileUtil xl=new ExcelfileUtil(inputpath);
		//count no of rows in a sheet
		int rc=xl.rowCount("Supplier");
		Reporter.log("no of rows are:"+rc,true);
		for(int i=1;i<=rc;i++) {
			//read all cell from supplier sheet
			String sname=xl.getCellData("Supplier", i,0);
			String address=xl.getCellData("Supplier", i, 1);
			String city=xl.getCellData("Supplier", i, 2);
			String country=xl.getCellData("Supplier", i, 3);
			String cperson=xl.getCellData("Supplier", i, 4);
			String pnumber=xl.getCellData("Supplier", i, 5);
			String mail=xl.getCellData("Supplier", i, 6);
			String mnumber=xl.getCellData("Supplier", i, 7);
			String notes=xl.getCellData("Supplier", i, 8);
			String Results=ERP_Function.verifySuppler(sname, address, city, country, cperson, pnumber, mail, mnumber, notes);
			Reporter.log(Results,true);
			xl.setCellData("Supplier", i, 9, Results, outpath);
		}
	}
	@AfterTest
	public void tearDown()throws Throwable
	{
		System.out.println("DataDriven.tearDown()");
		//	System.out.println(driver);
		ERP_Function.driver.close();
	}
}
