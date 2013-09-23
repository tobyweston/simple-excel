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

import bad.robot.excel.cell.BooleanCell;
import bad.robot.excel.cell.DoubleCell;
import bad.robot.excel.cell.StringCell;
import org.hamcrest.StringDescription;
import org.junit.Test;

import static bad.robot.excel.matchers.CellMatcher.equalTo;
import static bad.robot.excel.matchers.StubCell.createCell;
import static org.hamcrest.Matchers.is;
import static org.junit.Assert.assertThat;

public class CellMatcherTest {

    private final StringDescription description = new StringDescription();

    @Test
    public void usageExample() {
        assertThat(createCell("Text"), equalTo(new StringCell("Text")));
    }

    @Test
    public void matches() {
        assertThat(equalTo(new DoubleCell(0.44444d)).matches(createCell(0.44444d)), is(true));
    }

    @Test
    public void doesNotMatch() {
        assertThat(equalTo(new DoubleCell(0.44444d)).matches(createCell(0.4444d)), is(false));
    }

    @Test
    public void description() {
        equalTo(new DoubleCell(0.44444d)).describeTo(description);
        assertThat(description.toString(), is("<0.44444D>"));
    }

    @Test
    public void mismatch() {
        ((CellMatcher) equalTo(new BooleanCell(false))).matchesSafely(createCell(true), description);
        assertThat(description.toString(), is("cell at \"A1\" contained <TRUE> expected <FALSE>"));
    }

}
