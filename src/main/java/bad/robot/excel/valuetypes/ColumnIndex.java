package bad.robot.excel.valuetypes;

import bad.robot.AbstractValueType;

public class ColumnIndex extends AbstractValueType<Integer> {

    public static ColumnIndex column(Integer value) {
        return new ColumnIndex(value);
    }

    private ColumnIndex(Integer value) {
        super(value);
    }
}
