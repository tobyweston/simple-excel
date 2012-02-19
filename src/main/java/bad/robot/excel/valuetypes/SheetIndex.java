package bad.robot.excel.valuetypes;

import bad.robot.AbstractValueType;

public class SheetIndex extends AbstractValueType<Integer> {

    public static SheetIndex sheet(Integer value) {
        return new SheetIndex(value);
    }

    private SheetIndex(Integer value) {
        super(value);
    }
}
