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

package bad.robot.excel.column;

import org.junit.Test;

import static bad.robot.excel.column.ExcelColumnIndex.*;
import static org.hamcrest.Matchers.is;
import static org.junit.Assert.assertThat;

public class ExcelColumnIndexTest {

    @Test
    public void convertToIndex() {
        assertThat(getColumn("A").ordinal(), is(0));
        assertThat(getColumn("B").ordinal(), is(1));
        assertThat(getColumn("C").ordinal(), is(2));
        assertThat(getColumn("X").ordinal(), is(23));
        assertThat(getColumn("Y").ordinal(), is(24));
        assertThat(getColumn("Z").ordinal(), is(25));
    }

    @Test
    public void convertRowCoordinateToExcelRowForExtendedAlphabet() {
        assertThat(getColumn("AA").ordinal(),is(26));
        assertThat(getColumn("ZZ").ordinal(), is(701));
    }
}
