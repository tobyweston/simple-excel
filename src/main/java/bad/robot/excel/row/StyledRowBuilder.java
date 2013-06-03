package bad.robot.excel.row;

import bad.robot.excel.column.ColumnIndex;
import bad.robot.excel.style.Style;

import java.util.Date;

public interface StyledRowBuilder {
    StyledRowBuilder withBlank(ColumnIndex index, Style style);

    StyledRowBuilder withString(ColumnIndex index, String text, Style style);

    StyledRowBuilder withDouble(ColumnIndex index, Double value, Style style);

    StyledRowBuilder withInteger(ColumnIndex index, Integer value, Style style);

    StyledRowBuilder withDate(ColumnIndex index, Date date, Style style);

    StyledRowBuilder withFormula(ColumnIndex index, String formula, Style style);
}
