package bad.robot.excel.valuetypes;

import bad.robot.AbstractValueType;

public class CellIndex extends AbstractValueType<Integer> {

    private CellIndex(Integer value) {
        super(value);
    }

    public static CellIndex cellIndex(Integer value) {
        return new CellIndex(value);
    }

}
