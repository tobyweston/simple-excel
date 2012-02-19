package bad.robot.excel;

import bad.robot.AbstractValueType;

public class Formula extends AbstractValueType<String> {

    public static Formula formula(String formula) {
        return new Formula(formula);
    }

    private Formula(String formula) {
        super(formula);
    }
}
