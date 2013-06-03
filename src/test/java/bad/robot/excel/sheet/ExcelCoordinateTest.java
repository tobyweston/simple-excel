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

package bad.robot.excel.sheet;

import org.junit.Test;

import static org.hamcrest.Matchers.is;
import static org.junit.Assert.assertThat;

public class ExcelCoordinateTest {

    @Test
    public void convertRowCoordinateToExcelRowForBasicAlphabet() {
        assertThat(asColumn(0), is("A"));
        assertThat(asColumn(1), is("B"));
        assertThat(asColumn(2), is("C"));
        assertThat(asColumn(3), is("D"));
        assertThat(asColumn(4), is("E"));
        assertThat(asColumn(5), is("F"));
        assertThat(asColumn(6), is("G"));
        assertThat(asColumn(7), is("H"));
        assertThat(asColumn(8), is("I"));
        assertThat(asColumn(9), is("J"));
        assertThat(asColumn(10), is("K"));
        assertThat(asColumn(11), is("L"));
        assertThat(asColumn(12), is("M"));
        assertThat(asColumn(13), is("N"));
        assertThat(asColumn(14), is("O"));
        assertThat(asColumn(15), is("P"));
        assertThat(asColumn(16), is("Q"));
        assertThat(asColumn(17), is("R"));
        assertThat(asColumn(18), is("S"));
        assertThat(asColumn(19), is("T"));
        assertThat(asColumn(20), is("U"));
        assertThat(asColumn(21), is("V"));
        assertThat(asColumn(22), is("W"));
        assertThat(asColumn(23), is("X"));
        assertThat(asColumn(24), is("Y"));
        assertThat(asColumn(25), is("Z"));
    }

    @Test
    public void convertRowCoordinateToExcelRowForExtendedAlphabet() {
        assertThat(asColumn(0), is("A"));
        assertThat(asColumn(25), is("Z"));
        assertThat(asColumn(26), is("AA"));
        assertThat(asColumn(701), is("ZZ"));
        assertThat(asColumn(702), is("AAA"));
        assertThat(asColumn(703), is("AAB"));
        assertThat(asColumn(18277), is("ZZZ"));
        assertThat(asColumn(18278), is("AAAA"));
    }

    public String asColumn(int column) {
        String converted = "";
        while (column >= 0) {
            int remainder = column % 26;
            converted = (char) (remainder + 'A') + converted;
            column = (column / 26) - 1;
        }
        return converted;
    }

}
