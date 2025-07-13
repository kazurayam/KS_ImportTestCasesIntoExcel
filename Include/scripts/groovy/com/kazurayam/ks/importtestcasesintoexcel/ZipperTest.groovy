package com.kazurayam.ks.importtestcasesintoexcel

import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static org.junit.Assert.assertTrue

import com.kms.katalon.core.testcase.TestCase

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

import org.apache.commons.io.FileUtils
import org.junit.Before
import org.junit.BeforeClass
import org.junit.Test
import org.junit.runner.RunWith
import org.junit.runners.JUnit4

@RunWith(JUnit4.class)
public class ZipperTest {

	@BeforeClass
	static void beforeAll() {
		Path testOutput = Paths.get("build/tmp/testOutput")
		if (Files.exists(testOutput)) {
			FileUtils.deleteDirectory(testOutput.toFile())
		}
	}

	@Before
	void setup() {
		println "---------"
	}

	@Test
	void test_All() {
		String xlsx = "build/tmp/testOutput/test_All.xlsx"
		Path xlsxPath = Paths.get(xlsx)
		Zipper zipper = new Zipper(xlsx)
		zipper.execute()
		assertTrue(Files.exists(xlsxPath))
	}

	@Test
	void test_foo() {
		String xlsx = "build/tmp/testOutput/test_foo.xlsx"
		Path xlsxPath = Paths.get(xlsx)
		List<TestCase> includes = [ findTestCase("TC1") ]
		Zipper zipper = new Zipper(includes, xlsx)
		zipper.execute()
		assertTrue(Files.exists(xlsxPath))
	}
}
