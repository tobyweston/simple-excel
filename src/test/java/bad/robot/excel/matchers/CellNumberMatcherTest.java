/*
 * Copyright (c) 2012-2013, bad robot (london) ltd.
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

import org.apache.poi.ss.usermodel.Row;
import org.hamcrest.Description;
import org.hamcrest.StringDescription;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.WorkbookResource.firstRowOf;
import static bad.robot.excel.matchers.CellNumberMatcher.hasSameNumberOfCellsAs;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.is;
import static org.hamcrest.Matchers.not;

public class CellNumberMatcherTest {

    private Row rowWithThreeCells;
    private Row rowWithTwoCells;

    @Before
    public void loadWorkbookAndSheets() throws IOException {
        rowWithThreeCells = firstRowOf("rowWithThreeCells.xls");
        rowWithTwoCells = firstRowOf("rowWithTwoCells.xls");
    }

    @Test
    public void exampleUsages() throws Exception {
        assertThat(rowWithThreeCells, hasSameNumberOfCellsAs(rowWithThreeCells));
        assertThat(rowWithTwoCells, not(hasSameNumberOfCellsAs(rowWithThreeCells)));
    }

    @Test
    public void matches() {
        assertThat(hasSameNumberOfCellsAs(rowWithThreeCells).matches(rowWithThreeCells), is(true));
    }

    @Test
    public void doesNotMatch() {
        assertThat(hasSameNumberOfCellsAs(rowWithThreeCells).matches(rowWithTwoCells), is(false));
    }

    @Test
    public void description() {
        Description description = new StringDescription();
        hasSameNumberOfCellsAs(rowWithThreeCells).describeTo(description);
        assertThat(description.toString(), is("<3> cell(s) on row <1> sheet \"Sheet1\""));
    }

    @Test
    public void mismatch() {
        Description description = new StringDescription();
        hasSameNumberOfCellsAs(rowWithThreeCells).matchesSafely(rowWithTwoCells, description);
        assertThat(description.toString(), is("got <2> cell(s) on row <1> expected <3> sheet \"Sheet1\""));
    }

    @Test
    public void differentNumberOfPhysicalCells() throws IOException {
        Row rowWithThreePhysicalCells = firstRowOf("rowWithThreePhysicalCells.xls");
        Row rowWithThreeCellsButOnlyTwoPhysicalCells = firstRowOf("rowWithThreeCellsButOnlyTwoPhysicalCells.xls");
        Description description = new StringDescription();
        hasSameNumberOfCellsAs(rowWithThreeCellsButOnlyTwoPhysicalCells).matchesSafely(rowWithThreePhysicalCells, description);
        assertThat(description.toString(), is("got <3> cell(s) containing value on row <1> expected <2> sheet \"Sheet1\""));
    }

}
