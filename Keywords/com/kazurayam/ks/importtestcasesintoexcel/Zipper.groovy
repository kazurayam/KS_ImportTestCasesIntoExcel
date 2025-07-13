package com.kazurayam.ks.importtestcasesintoexcel

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths
import java.nio.charset.StandardCharsets

import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook

import com.kazurayam.ant.DirectoryScanner
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.configuration.RunConfiguration

public class Zipper {

	private Path baseDir = Paths.get(RunConfiguration.getProjectDir()).toAbsolutePath()

	private List<String> includePatterns

	private String xlsx
	
	private static final int SHEET_NAME_MAX_LENGTH = 29

	Zipper(String xlsx) {
		init(["**"], xlsx)
	}

	Zipper(List<TestCase> testCases, String xlsx) {
		List<String> list = new ArrayList<>()
		for (TestCase tc in testCases) {
			String tcId = tc.getTestCaseId().replace("Test Cases/", "")
			list.add(tcId)
		}
		init(list, xlsx)
	}

	private init(List<String> testCaseIds, String xlsx) {
		this.includePatterns = new ArrayList<>()
		for (String s in testCaseIds) {
			if (s.length() > 0) {
				this.includePatterns.add("Scripts/" + s + "/**/*.groovy")
			}
		}
		this.xlsx = xlsx
	}

	void execute() {
		HSSFWorkbook workbook = new HSSFWorkbook()
		//
		DirectoryScanner ds = new DirectoryScanner()
		ds.setBasedir(this.baseDir.toString())
		ds.setIncludes(this.includePatterns.toArray(new String[0]))
		ds.scan()
		String[] groovyFiles = ds.getIncludedFiles()
		println "baseDir: " + this.baseDir
		println "includePatterns: " + this.includePatterns
		println "number of files: " + groovyFiles.length
		for (int i = 0; i < groovyFiles.length; i++) {
			Path p = baseDir.resolve(Paths.get(groovyFiles[i]))
			println p
			HSSFSheet spreadSheet = workbook.createSheet(toSheetName(p))
			importScriptIntoSheet(p, spreadSheet)
		}
		//
		writeWorkbookIntoFile(workbook, xlsx)
	}

	private void importScriptIntoSheet(Path p, HSSFSheet spreadSheet) {
		BufferedReader br =
				new BufferedReader(new InputStreamReader(p.newInputStream(), StandardCharsets.UTF_8))
		String line
		int ln = 0
		while((line = br.readLine()) != null) {
			HSSFRow row = spreadSheet.createRow((short) ln)
			HSSFCell cell0 = row.createCell(0)
			cell0.setCellValue(line)
			ln++
		}
	}

	private void writeWorkbookIntoFile(HSSFWorkbook workbook, String excelFilePath) {
		ByteArrayOutputStream baOut = new ByteArrayOutputStream()
		workbook.write(baOut)
		//
		Path xlsx = Paths.get(excelFilePath)
		Files.createDirectories(xlsx.getParent())
		FileOutputStream fOut = new FileOutputStream(xlsx.toFile())
		baOut.writeTo(fOut)
		baOut.close()
		fOut.close()
	}

	private String toSheetName(Path groovyFile) {
		Path relativePath = this.baseDir.relativize(groovyFile)
		Path parentDir = relativePath.getParent()
		String s = parentDir.toString().replaceAll("Scripts/", "")
		if (s.length() > SHEET_NAME_MAX_LENGTH) {
			s = parentDir.getFileName().toString().replaceAll("Scripts/", "")
		}
		return s.replaceAll("/","／")
	}
}
