/*
 * Copyright (c) 2012-2019, bad robot (london) ltd.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package bad.robot.excel.matchers;

import org.apache.poi.ss.usermodel.Sheet;
import org.hamcrest.Description;
import org.hamcrest.StringDescription;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.WorkbookResource.firstSheetOf;
import static bad.robot.excel.matchers.MergedRegionsMatcher.hasSameMergedRegions;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.is;

public class MergedRegionsMatcherTest {

	private Sheet sheetWithMergedRegions;
	private Sheet sheetWithAltMergedRegion;
	private Sheet sheetWithTwoMergedRegion;
	private Sheet sheetWithMergedRegionsAltText;

	@Before
	public void loadWorkbookAndSheets() throws IOException {
		sheetWithMergedRegions = firstSheetOf("mergedRegionSheet.xlsx");
		sheetWithAltMergedRegion = firstSheetOf("mergedRegionWithAltRegion.xlsx");
		sheetWithTwoMergedRegion = firstSheetOf("mergedRegionWithTwoRegions.xlsx");
		sheetWithMergedRegionsAltText = firstSheetOf("mergedRegionWithAltText.xlsx");
	}

	@Test
	public void exampleUsages() {
		assertThat(sheetWithMergedRegions, hasSameMergedRegions(sheetWithMergedRegions));
	}

	@Test
	public void matches() {
		assertThat(hasSameMergedRegions(sheetWithMergedRegions).matches(sheetWithMergedRegions), is(true));
	}
	
	@Test
	public void matchesIgnoreCellContents() {
		// cells contain different text values but are both merged cells
		assertThat(hasSameMergedRegions(sheetWithMergedRegions).matches(sheetWithMergedRegionsAltText), is(true));
	}

	@Test
	public void doesNotMatch() {
		assertThat(hasSameMergedRegions(sheetWithMergedRegions).matches(sheetWithAltMergedRegion), is(false));
	}

	@Test
	public void description() {
		Description description = new StringDescription();
		hasSameMergedRegions(sheetWithMergedRegions).describeTo(description);
		assertThat(description.toString(), is("\"A1:E1\" as a merged region(s) in sheet \"Sheet1\""));
	}

	@Test
	public void descriptionAltRegion() {
		Description description = new StringDescription();
		hasSameMergedRegions(sheetWithAltMergedRegion).describeTo(description);
		assertThat(description.toString(), is("\"A2:E2\" as a merged region(s) in sheet \"Sheet1\""));
	}

	@Test
	public void descriptionMultipleMismatches() {
		Description description = new StringDescription();
		hasSameMergedRegions(sheetWithTwoMergedRegion).describeTo(description);
		assertThat(description.toString(), is("\"A2:E2, A4:E4\" as a merged region(s) in sheet \"Sheet1\""));
	}

	@Test
	public void mismatch() {
		Description description = new StringDescription();
		hasSameMergedRegions(sheetWithMergedRegions).matchesSafely(sheetWithAltMergedRegion, description);
		assertThat(description.toString(), is("got \"A2:E2\" as a merged region(s) in sheet \"Sheet1\" expected \"A1:E1\""));
	}

}