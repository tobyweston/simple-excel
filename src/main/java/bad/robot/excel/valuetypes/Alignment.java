package bad.robot.excel.valuetypes;

import bad.robot.AbstractValueType;
import bad.robot.excel.AlignmentStyle;

public class Alignment extends AbstractValueType<AlignmentStyle> {

    public static Alignment alignment(AlignmentStyle value) {
        return new Alignment(value);
    }

    public Alignment(AlignmentStyle value) {
        super(value);
    }
}
