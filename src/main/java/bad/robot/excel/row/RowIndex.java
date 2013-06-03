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

package bad.robot.excel.row;

import bad.robot.excel.AbstractValueType;

public class RowIndex extends AbstractValueType<Integer> {

    public static RowIndex row(Integer value) {
        if (value <= 0)
            throw new IllegalArgumentException("row indices start at 1");
        return new RowIndex(value - 1);
    }

    private RowIndex(Integer value) {
        super(value);
    }
}
