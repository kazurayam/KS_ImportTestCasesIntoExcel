package com.kazurayam.ks.importtestcasesintoexcel

import static org.junit.Assert.assertEquals
import static org.junit.Assert.assertNotEquals
import static org.junit.Assert.assertTrue
import static org.junit.Assert.fail

import org.junit.Test
import org.junit.runner.RunWith
import org.junit.runners.JUnit4

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

@RunWith(JUnit4.class)
public class ZipperTest {
	
	@Test
	void test_smoke() {
		String tcId = "TC1"
		String xlsx = "build/tmp/testOutput/out.xlsx"
		Path xlsxPath = Paths.get(xlsx)
		Zipper zipper = new Zipper(tcId, xlsx)
		zipper.execute()
		assertTrue(Files.exists(xlsxPath))
	}
}
