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

package bad.robot.excel.valuetypes;

import org.junit.Test;

import static bad.robot.excel.valuetypes.ExcelColumn.*;
import static org.hamcrest.Matchers.is;
import static org.junit.Assert.assertThat;

public class ExcelColumnTest {

    @Test
    public void convertToIndex() {
        assertThat(A.ordinal(), is(0));
        assertThat(B.ordinal(), is(1));
        assertThat(C.ordinal(), is(2));
        assertThat(X.ordinal(), is(23));
        assertThat(Y.ordinal(), is(24));
        assertThat(Z.ordinal(), is(25));
    }

    @Test
    public void convertRowCoordinateToExcelRowForExtendedAlphabet() {
        assertThat(AA.ordinal(),is(26));
        assertThat(ZZ.ordinal(), is(701));
        assertThat(AAA.ordinal(), is(702));
        assertThat(AAB.ordinal(), is(703));
        assertThat(ZZZ.ordinal(), is(18277));
        assertThat(AAAA.ordinal(), is(18278));
    }
}
