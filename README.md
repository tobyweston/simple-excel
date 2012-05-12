# bad.robot.simple-excel

Basically wrapping the POI project with Java builders to modify simple sheets quickly and easily.

## simple modification of an Excel sheet 

Add styles, formula and content to cells programmatically via a simple DSL

    @Test
    public void shouldReplaceCellsInComplicatedAlternateSyntaxExample() throws IOException {
        HSSFWorkbook workbook = getWorkbook("shouldReplaceCellsInComplicatedExampleTemplate.xls");
        new PoiWorkbookMutator(workbook)
                .replaceCell(coordinate(C, 5), "Very")
                .replaceCell(coordinate(D, 11), "Complicated")
                .replaceCell(coordinate(G, 3), "Example")
                .replaceCell(coordinate(H, 7), "Of")
                .replaceCell(coordinate(J, 10), "Templated")
                .replaceCell(coordinate(M, 15), "Spreadsheet");

        assertThat(workbook, is(sameWorkBook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplateExpected.xls"))));
    }

A break in the matcher would show something like

    java.lang.AssertionError:
    Expected: value 'Templated' at Cell J10
         got: value 'templated'


## more.tools

For more tools, see [robotooling.com](http://www.robotooling.com)