# bad.robot.simple-excel

Basically wrapping the POI project with Java builders to modify simple sheets quickly and easily.

## simple modification of an Excel sheet 

Add styles, formula and content to cells programatically via a simple DSL

    @Test
    public void shouldReplaceCellsInComplicatedExample() throws IOException {
        HSSFWorkbook workbook = getWorkbook("shouldReplaceCellsInComplicatedExampleTemplate.xls");
        new PoiWorkbookMutator(workbook)
                .replaceCell(coordinate(column(2), row(4)), "Very")
                .replaceCell(coordinate(column(3), row(10)), "Complicated")
                .replaceCell(coordinate(column(6), row(2)), "Example")
                .replaceCell(coordinate(column(7), row(6)), "Of")
                .replaceCell(coordinate(column(9), row(9)), "Templated")
                .replaceCell(coordinate(column(12), row(14)), "Spreadsheet");

        assertThat(workbook, is(sameWorkBook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplateExpected.xls"))));
    }

## more.tools

For more tools, see [robotooling.com](http://www.robotooling.com)