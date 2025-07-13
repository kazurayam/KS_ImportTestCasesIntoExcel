package com.kazurayam.ks.importtestcasesintoexcel

import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static org.assertj.core.api.InstanceOfAssertFactories.THROWABLE

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRichTextString

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

public class Zipper {

	private String testCaseId

	private String excelFilePath

	Zipper(String testCaseId, String excelFilePath) {
		this.testCaseId = testCaseId
		this.excelFilePath = excelFilePath
	}

	void execute() {
		Path testCaseScript = Paths.get(findTestCase(testCaseId).getGroovyScriptPath())
		if (!Files.exists(testCaseScript)) {
			throw new IOException(testCaseScript + " is not present")
		}
		Path xlsx = Paths.get(excelFilePath)
		Files.createDirectories(xlsx.getParent())
		
		//
		HSSFWorkbook workBook = new HSSFWorkbook()
		
		//
		HSSFSheet spreadSheet = workBook.createSheet(testCaseId)
		
		// 
		HSSFRow row = spreadSheet.createRow((short) 0)
		
		//
		HSSFCell cell0 = row.createCell(0)
		cell0.setCellValue(new HSSFRichTextString("Hello"))
		
		HSSFCell cell1 = row.createCell(1)
		cell1.setCellValue(new HSSFRichTextString("World"))
		
		//
		FileOutputStream fout = new FileOutputStream(xlsx.toFile())
		
		//
		ByteArrayOutputStream baOut = new ByteArrayOutputStream()
		workBook.write(baOut)

		FileOutputStream fOut = new FileOutputStream(xlsx.toFile())
		baOut.writeTo(fOut)
		baOut.close()
		fOut.close()
	}
}
