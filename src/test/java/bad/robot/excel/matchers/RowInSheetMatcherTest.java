/*
 * Copyright (c) 2012-2013, bad robot (london) ltd.
 *
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */

package bad.robot.excel.matchers;

import org.apache.poi.ss.usermodel.Sheet;
import org.hamcrest.StringDescription;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.WorkbookResource.firstRowOf;
import static bad.robot.excel.WorkbookResource.firstSheetOf;
import static bad.robot.excel.matchers.RowInSheetMatcher.hasSameRow;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.is;
import static org.hamcrest.Matchers.not;

public class RowInSheetMatcherTest {

    private final StringDescription description = new StringDescription();

    private Sheet sheetWithThreeCells;
    private Sheet sheetWithThreeCellsAlternativeValues;
    private Sheet sheetWithTwoCells;

    @Before
    public void loadWorkbooks() throws IOException {
        sheetWithThreeCells = firstSheetOf("rowWithThreeCells.xls");
        sheetWithThreeCellsAlternativeValues = firstSheetOf("rowWithThreeCellsAlternativeValues.xls");
        sheetWithTwoCells = firstSheetOf("rowWithTwoCells.xls");
    }

    @Test
    public void exampleUsage() throws IOException {
        assertThat(sheetWithTwoCells, hasSameRow(firstRowOf(sheetWithTwoCells)));
        assertThat(sheetWithThreeCells, not(hasSameRow(firstRowOf(sheetWithTwoCells))));
    }

    @Test
    public void matches() {
        assertThat(hasSameRow(firstRowOf(sheetWithThreeCells)).matches(sheetWithThreeCells), is(true));
    }

    @Test
    public void doesNotMatch() {
        assertThat(hasSameRow(firstRowOf(sheetWithThreeCells)).matches(sheetWithThreeCellsAlternativeValues), is(false));
    }

    @Test
    public void description() {
        hasSameRow(firstRowOf(sheetWithThreeCells)).describeTo(description);
        assertThat(description.toString(), is("equality of row <1>"));
    }

    @Test
    public void mismatch() {
        hasSameRow(firstRowOf(sheetWithThreeCells)).matchesSafely(sheetWithThreeCellsAlternativeValues, description);
        assertThat(description.toString(), is("cell at \"B1\" contained <3.14D> expected <\"C2, R1\">"));
    }
}
