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
import static bad.robot.excel.matchers.MergedRegionCountMatcher.hasSameNumberOfMergedRegions;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.is;
import static org.hamcrest.Matchers.not;

public class MergedRegionCountMatcherTest {

	private Sheet sheetWithMergedRegions;
	private Sheet emptySheet;

	@Before
	public void loadWorkbookAndSheets() throws IOException {
		sheetWithMergedRegions = firstSheetOf("mergedRegionSheet.xlsx");
		emptySheet = firstSheetOf("emptySheet.xlsx");
	}

	@Test
	public void exampleUsages() {
		assertThat(sheetWithMergedRegions, hasSameNumberOfMergedRegions(sheetWithMergedRegions));
		assertThat(emptySheet, not(hasSameNumberOfMergedRegions(sheetWithMergedRegions)));
	}

	@Test
	public void matches() {
		assertThat(hasSameNumberOfMergedRegions(sheetWithMergedRegions).matches(sheetWithMergedRegions), is(true));
	}

	@Test
	public void doesNotMatch() {
		assertThat(hasSameNumberOfMergedRegions(sheetWithMergedRegions).matches(emptySheet), is(false));
	}

	@Test
	public void description() {
		Description description = new StringDescription();
		hasSameNumberOfMergedRegions(sheetWithMergedRegions).describeTo(description);
		assertThat(description.toString(), is("<1> merged region(s) in sheet \"Sheet1\""));
	}

	@Test
	public void mismatch() {
		Description description = new StringDescription();
		hasSameNumberOfMergedRegions(sheetWithMergedRegions).matchesSafely(emptySheet, description);
		assertThat(description.toString(), is("got <0> merged region(s) in sheet \"Sheet1\" expected <1>"));
	}

}