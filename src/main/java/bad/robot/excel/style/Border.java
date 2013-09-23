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

package bad.robot.excel.style;

public class Border {

    private final TopBorder top;
    private final BottomBorder bottom;
    private final LeftBorder left;
    private final RightBorder right;

    private Border(TopBorder top, RightBorder right, BottomBorder bottom, LeftBorder left) {
        this.top = top;
        this.bottom = bottom;
        this.left = left;
        this.right = right;
    }

    public static Border border(TopBorder top, RightBorder right, BottomBorder bottom, LeftBorder left) {
        return new Border(top, right, bottom, left);
    }

    public BottomBorder getBottom() {
        return bottom;
    }

    public TopBorder getTop() {
        return top;
    }

    public LeftBorder getLeft() {
        return left;
    }

    public RightBorder getRight() {
        return right;
    }
}
