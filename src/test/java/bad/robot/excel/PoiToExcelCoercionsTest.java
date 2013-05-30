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

package bad.robot.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.hamcrest.Matcher;
import org.jmock.Expectations;
import org.jmock.Mockery;
import org.junit.Test;

import static bad.robot.excel.PoiToExcelCoercions.asExcelCoordinate;
import static bad.robot.excel.PoiToExcelCoercionsTest.PoiCoordinate.coordinate;
import static org.hamcrest.Matchers.is;
import static org.junit.Assert.assertThat;

public class PoiToExcelCoercionsTest {

    private final Mockery context = new Mockery();
    private final Cell cell = context.mock(Cell.class);

    @Test
    public void shouldConvertFromPoiToExcelCoordinates() {
        verify(coordinate(0, 0), is("A1"));
        verify(coordinate(1, 1), is("B2"));
        verify(coordinate(10, 5), is("K6"));
        verify(coordinate(26, 11), is("AA12"));
        verify(coordinate(69, 15), is("BR16"));
    }

    private void verify(PoiCoordinate coordinate, Matcher<String> matcher) {
        expectingCellAt(coordinate);
        assertThat(asExcelCoordinate(cell), matcher);
        context.assertIsSatisfied();
    }

    private void expectingCellAt(final PoiCoordinate coordinate) {
        context.checking(new Expectations() {{
            oneOf(cell).getColumnIndex(); will(returnValue(coordinate.column));
            oneOf(cell).getRowIndex(); will(returnValue(coordinate.row));
        }});
    }

    static class PoiCoordinate {
        private final int column;
        private final int row;

        private PoiCoordinate(int column, int row) {
            this.column = column;
            this.row = row;
        }

        static PoiCoordinate coordinate(int column, int row) {
            return new PoiCoordinate(column, row);
        }
    }
}
