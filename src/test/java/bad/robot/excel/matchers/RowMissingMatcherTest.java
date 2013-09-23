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
import org.hamcrest.StringDescription;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.WorkbookResource.thirdRowOf;
import static bad.robot.excel.matchers.RowMissingMatcher.rowIsPresent;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.is;
import static org.hamcrest.Matchers.not;

public class RowMissingMatcherTest {

    private final StringDescription description = new StringDescription();

    private Row thirdRow;
    private Row sheetWithThreeRows;

    @Before
    public void loadWorkbooks() throws IOException {
        thirdRow = thirdRowOf("sheetWithTwoRows.xls");
        sheetWithThreeRows = thirdRowOf("sheetWithThreeRows.xls");
    }

    @Test
    public void usageExample() throws Exception {
        assertThat(thirdRow, not(rowIsPresent(sheetWithThreeRows)));
    }

    @Test
    public void matches() {
        assertThat(rowIsPresent(sheetWithThreeRows).matches(sheetWithThreeRows), is(true));
    }

    @Test
    public void doesNotMatch() {
        assertThat(rowIsPresent(sheetWithThreeRows).matches(thirdRow), is(false));
    }

    @Test
    public void describe() {
        rowIsPresent(sheetWithThreeRows).describeTo(description);
        assertThat(description.toString(), is("row <3> to be present"));
    }

    @Test
    public void mismatch() {
        rowIsPresent(sheetWithThreeRows).matchesSafely(thirdRow, description);
        assertThat(description.toString(), is("row <3> is missing"));
    }

}
