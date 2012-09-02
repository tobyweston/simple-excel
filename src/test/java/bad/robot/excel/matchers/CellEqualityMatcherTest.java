/*
 * Copyright (c) 2012, bad robot (london) ltd.
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
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.WorkbookResource.firstRowOf;
import static bad.robot.excel.matchers.CellEqualityMatcher.hasSameCellsAs;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.is;
import static org.hamcrest.Matchers.not;

public class CellEqualityMatcherTest {

    private StringDescription description = new StringDescription();

    private Row rowWithThreeCells;
    private Row rowWithThreeCellsAlternateValues;

    @Before
    public void loadWorkbookAndSheets() throws IOException {
        rowWithThreeCells = firstRowOf("rowWithThreeCells.xls");
        rowWithThreeCellsAlternateValues = firstRowOf("rowWithThreeCellsAlternativeValues.xls");
    }

    @Test
    public void exampleUsage() {
        assertThat(rowWithThreeCells, hasSameCellsAs(rowWithThreeCells));
        assertThat(rowWithThreeCellsAlternateValues, not(hasSameCellsAs(rowWithThreeCells)));
    }

    @Test
    public void matches() {
        assertThat(hasSameCellsAs(rowWithThreeCells).matches(rowWithThreeCells), is(true));
    }

    @Test
    public void doesNotMatch() {
        assertThat(hasSameCellsAs(rowWithThreeCellsAlternateValues).matches(rowWithThreeCells), is(false));
    }

    @Test
    public void description() {
        hasSameCellsAs(rowWithThreeCells).describeTo(description);
        assertThat(description.toString(), is("equality of all cells of row <1>"));
    }

    @Test
    public void mismatch() {
        hasSameCellsAs(rowWithThreeCells).matchesSafely(rowWithThreeCellsAlternateValues, description);
        assertThat(description.toString(), is("cell at \"B1\" contained <3.14D> expected <\"Cell 2, Row 1\">"));
    }
}
