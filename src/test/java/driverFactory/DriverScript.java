package driverFactory;

import org.openqa.selenium.WebDriver;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import commonFunctions.FunctionLibrary;
import utilities.ExcelFileUtil;

public class DriverScript {

	WebDriver driver;

	String inputpath = "./FileInput/DataEngine.xlsx";

	String outputpath = "./FileOutput/HybridResults.xlsx";

	ExtentReports report;

	ExtentTest logger;

	String TestCases = "MasterTestCases";

	public void startTest() throws Throwable {

		String Module_Status;

		//Create Object for ExcelFileUtil class
		ExcelFileUtil xl = new ExcelFileUtil(inputpath);

		//Iterate all TestCases in TestCases
		for(int i=1;i<=xl.rowCount(TestCases);i++) {

			if(xl.getCellData(TestCases, i, 2).equalsIgnoreCase("Y"))
			{
				//read corresponding sheet or TestCases
				String TCModule = xl.getCellData(TestCases, i, 1);

				//Define path of HTML Reports
				report = new ExtentReports("./target/ExtentReports/"+TCModule+FunctionLibrary.generateDate()+".html");

				logger = report.startTest(TCModule);

				logger.assignAuthor("Tarak");

				//Iterate all row in a TCModule sheet
				for(int j=1;j<=xl.rowCount(TCModule);j++) {

					//read all cell from TCModule sheet
					String Description = xl.getCellData(TCModule, j, 0);
					String Object_Type = xl.getCellData(TCModule, j, 1);					
					String LName = xl.getCellData(TCModule, j, 2);					
					String LValue = xl.getCellData(TCModule, j, 3);					
					String TestData = xl.getCellData(TCModule, j, 4);

					try {

						if(Object_Type.equalsIgnoreCase("startBrowser")) {

							driver = FunctionLibrary.startBrowser();
							logger.log(LogStatus.INFO, Description);
						}

						if(Object_Type.equalsIgnoreCase("openUrl")) {

							FunctionLibrary.openUrl();
							logger.log(LogStatus.INFO, Description);
						}

						if(Object_Type.equalsIgnoreCase("waitForElement")) {

							FunctionLibrary.waitForElement(LName, LValue, TestData);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("typeAction")) {

							FunctionLibrary.typeAction(LName, LValue, TestData);
							logger.log(LogStatus.INFO, Description);
						}

						if(Object_Type.equalsIgnoreCase("clickAction")) {

							FunctionLibrary.clickAction(LName, LValue);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("validateTitle")) {

							FunctionLibrary.validateTitle(TestData);
							logger.log(LogStatus.INFO, Description);
						}

						if(Object_Type.equalsIgnoreCase("closeBrowser")) {

							FunctionLibrary.closeBrowser();
							logger.log(LogStatus.INFO, Description);
						}
						
						if(Object_Type.equalsIgnoreCase("mouseclick")) {
							
							FunctionLibrary.mouseclick();
							logger.log(LogStatus.INFO, Description);
						}
						
						if(Object_Type.equalsIgnoreCase("categoryTable")) {
							
							FunctionLibrary.categoryTable(TestData);
							logger.log(LogStatus.INFO, Description);
						}
						
						if(Object_Type.equalsIgnoreCase("dropDownAction")) {
							
							FunctionLibrary.dropDownAction(LName, LValue, TestData);
							logger.log(LogStatus.INFO, Description);
						}
						
						if(Object_Type.equalsIgnoreCase("captureStock")){
							
							FunctionLibrary.captureStock(LName, LValue);
							logger.log(LogStatus.INFO, Description);
						}
						
						if(Object_Type.equalsIgnoreCase("stockTable")) {
							
							FunctionLibrary.stockTable();
							logger.log(LogStatus.INFO, Description);
						}
						
						if(Object_Type.equalsIgnoreCase("captureSupplier")) {
							
							FunctionLibrary.captureSupplier(LName, LValue);
							logger.log(LogStatus.INFO, Description);
						}
						
						if(Object_Type.equalsIgnoreCase("supplierTable")) {
							
							FunctionLibrary.supplierTable();
							logger.log(LogStatus.INFO, Description);
						}
						
						if(Object_Type.equalsIgnoreCase("captureCustomer")) {
							
							FunctionLibrary.captureCustomer(LName, LValue);
							logger.log(LogStatus.INFO, Description);
						}
						
						if(Object_Type.equalsIgnoreCase("CustomerTable")) {
							
							FunctionLibrary.CustomerTable();
							logger.log(LogStatus.INFO, Description);
						}
						
						
						

						//write as pass into TCModule sheet in status cell
						xl.setCellData(TCModule, j, 5, "PASS", outputpath);
						logger.log(LogStatus.PASS, Description);
						Module_Status="True";
					}
					catch (Exception e) 
					{

						System.out.println(e.getMessage());

						//Write as a "FAIL" into TCModule Sheet in Status Cell
						xl.setCellData(TCModule, j, 5, "Fail", outputpath);
						logger.log(LogStatus.FAIL, Description);
						Module_Status = "False";
					}

					//Updating MasterTestSheet Status column PASS/FAIL
					if(Module_Status.equalsIgnoreCase("True")) {

						//Write as PASS into MASTERTestSheet Status
						xl.setCellData(TestCases, i, 3, "Pass", outputpath);
					}
					else {

						//Write as FAIL into MASTERTestSheet Status
						xl.setCellData(TestCases, i, 3, "Fail", outputpath);
					}

					report.endTest(logger);

					report.flush();
				}
			}
			else {
				//Write as Block for TestCases flag to N in TestCases Sheet
				xl.setCellData(TestCases, i, 3, "Blocked", outputpath);
			}

		}
	}	

}
