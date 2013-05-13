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

import org.apache.poi.ss.usermodel.Row;
import org.hamcrest.StringDescription;
import org.junit.Before;
import org.junit.Ignore;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.WorkbookResource.*;
import static bad.robot.excel.matchers.CellsMatcher.hasSameCellsAs;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.*;
import static org.hamcrest.core.AllOf.allOf;

public class CellsMatcherTest {

    private StringDescription description = new StringDescription();

    private Row firstRow;
    private Row secondRow;
    private Row thirdRow;
    private Row firstRowWithAlternateValues;
    private Row secondRowWithAlternateValues;
    private Row thirdRowWithAlternateValues;

    @Before
    public void loadWorkbookAndSheets() throws IOException {
        firstRow = firstRowOf("rowWithThreeCells.xls");
        secondRow = secondRowOf("rowWithThreeCells.xls");
        thirdRow = thirdRowOf("rowWithThreeCells.xls");
        firstRowWithAlternateValues = firstRowOf("rowWithThreeCellsAlternativeValues.xls");
        secondRowWithAlternateValues = secondRowOf("rowWithThreeCellsAlternativeValues.xls");
        thirdRowWithAlternateValues = thirdRowOf("rowWithThreeCellsAlternativeValues.xls");
    }

    @Test
    @Ignore
    public void exampleUsage() {
        assertThat(firstRow, hasSameCellsAs(firstRow));
        assertThat(firstRowWithAlternateValues, not(hasSameCellsAs(firstRow)));
        assertThat(secondRowWithAlternateValues, not(hasSameCellsAs(secondRow)));
        assertThat(thirdRowWithAlternateValues, not(hasSameCellsAs(thirdRow)));
    }

    @Test
    @Ignore
    public void matches() {
        assertThat(hasSameCellsAs(firstRow).matches(firstRow), is(true));
    }

    @Test
    @Ignore
    public void doesNotMatch() {
        assertThat(hasSameCellsAs(firstRowWithAlternateValues).matches(firstRow), is(false));
    }

    @Test
    @Ignore
    public void description() {
        hasSameCellsAs(firstRow).describeTo(description);
        assertThat(description.toString(), is("equality of all cells on row <1>"));
    }

    @Test
    @Ignore
    public void mismatch() {
        hasSameCellsAs(firstRow).matchesSafely(firstRowWithAlternateValues, description);
        assertThat(description.toString(), is("cell at \"B1\" contained <3.14D> expected <\"C2, R1\">"));
    }

    @Test
    @Ignore
    public void mismatchOnMissingCell() {
        hasSameCellsAs(secondRow).matchesSafely(secondRowWithAlternateValues, description);
        assertThat(description.toString(), is("cell at \"B2\" contained <nothing> expected <\"C2, R2\">"));
    }

    @Test
    public void mismatchMultipleValues() {
        hasSameCellsAs(thirdRow).matchesSafely(thirdRowWithAlternateValues, description);
        assertThat(description.toString(), allOf(
            containsString("cell at \"A3\" contained <Wed Feb 01 00:00:00 GMT 2012> expected <\"C1, R3\">,"),
            containsString("cell at \"B3\" contained <Formula:2+2> expected <\"C2, R3\">")
        ));
    }
}
