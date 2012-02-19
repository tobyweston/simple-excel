package bad.robot.excel.valuetypes;

import bad.robot.AbstractValueType;

public class RowIndex extends AbstractValueType<Integer> {

    public static RowIndex row(Integer value) {
        return new RowIndex(value);
    }

    private RowIndex(Integer value) {
        super(value);
    }
}
