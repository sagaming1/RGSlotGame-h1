import com.kms.katalon.core.configuration.RunConfiguration as RunConfiguration
import com.kms.katalon.core.testdata.reader.ExcelFactory as ExcelFactory
import org.apache.poi.xssf.usermodel.XSSFCell as XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow as XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet as XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook as XSSFWorkbook
import java.lang.String as String
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
//import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import internal.GlobalVariable as GlobalVariable
import com.kms.katalon.core.logging.KeywordLogger as KeywordLogger

String dirpath = RunConfiguration.getProjectDir()

KeywordLogger log = new KeywordLogger()

//third parameter means if you want the first row as your header or column name.
//In your case, it should be true.
FileInputStream file = new FileInputStream(new File(dirpath + '//Data//JDB-Data.xlsx'))

XSSFWorkbook workbook = new XSSFWorkbook(file)

XSSFSheet sheet = workbook.getSheetAt(0)

Object excelFile = ExcelFactory.getExcelDataWithDefaultSheet(dirpath + '//Data//JDB-Data.xlsx', 'sheet1', true)

//println(excelFile.getRowNumbers()) //Get total rows of the test data
//https://docs.katalon.com/javadoc/com/kms/katalon/core/testdata/reader/SheetPOI.html
WebUI.openBrowser('http://h1-game.rg-show.com/SlotGame/JDB')

String path = "//span[@class='txt-gameName']"
TestObject product = new TestObject()
product = WebUI.modifyObjectProperty(product, 'xpath', 'equals', path, true)
WebUI.waitForElementVisible(product, 1000)

int success = 0

int failure = 0

for (int i = 0; i < (sheet.size() - 1); i++) {
    //println(excelFile.internallyGetValue(0, i))
    if (WebUI.verifyTextPresent(excelFile.internallyGetValue(0, i), false, FailureHandling.OPTIONAL)) {
        println(excelFile.internallyGetValue(0, i) + '.......present')

        success = (success + 1)
    } else {
        println(excelFile.internallyGetValue(0, i) + '.......NOT present')

        log.logFailed(excelFile.internallyGetValue(0, i) + ' is not present')

        failure = (failure + 1)
    }
}

println((('成功上架Success:' + success) + ' 應上架未上架Failure:') + failure)

//關閉EXCEL********
file.close()

FileOutputStream outFile = new FileOutputStream(new File(dirpath + '//Data//JDB-Data.xlsx'))

workbook.write(outFile)

outFile.close()

