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

public class Zipper {

	private String baseDir = "Scripts"
	
	private String[] includes
	
	private String xlsx

	Zipper(String xlsx) {
		this(["**/*.groovy"], xlsx)
	}
	
	Zipper(String[] includes, String xlsx) {		
		this.includes = includes
		this.xlsx = xlsx
	}

	void execute() {
		HSSFWorkbook workbook = new HSSFWorkbook()
		//
		DirectoryScanner ds = new DirectoryScanner()
		ds.setBasedir(this.baseDir)
		ds.setIncludes(this.includes)
		ds.scan()
		String[] groovyFiles = ds.getIncludedFiles()
		println "baseDir: " + this.baseDir
		println "includes: " + this.includes
		println "number of files: " + groovyFiles.length
		for (int i = 0; i < groovyFiles.length; i++) {
			Path p = Paths.get(baseDir).resolve(Paths.get(groovyFiles[i]))
			println p
			HSSFSheet spreadSheet = workbook.createSheet(resolveTestCaseId(p))
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
	
	private String resolveTestCaseId(Path groovyFile) {
		return groovyFile.getParent().toString().replaceAll("Scripts/","").replaceAll("/","／")
	}
}
