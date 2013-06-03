package bad.robot.excel.row;

import bad.robot.excel.column.ColumnIndex;

import java.util.Date;

public interface RowBuilder {

    RowBuilder withBlank(ColumnIndex index);

    RowBuilder withString(ColumnIndex index, String text);

    RowBuilder withDouble(ColumnIndex index, Double value);

    RowBuilder withInteger(ColumnIndex index, Integer value);

    RowBuilder withDate(ColumnIndex index, Date date);

    RowBuilder withFormula(ColumnIndex index, String formula);

    Row build();
}
