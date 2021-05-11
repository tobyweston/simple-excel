# Simple-Excel
[![](https://travis-ci.org/tobyweston/simple-excel.png?branch=master)](https://travis-ci.org/tobyweston/simple-excel)
## Simply modify and diff Excel sheets in Java

Simple-Excel wraps the [Apache POI](https://poi.apache.org/) project with simple Java builders to modify sheets quickly and easily without all the boilerplate.

Use [Hamcrest](http://hamcrest.org/) `Matcher`s to compare workbooks and get fast feedback in tests. Comparing two sheets will compare the entire contents. You get a full report of the diff rather than just the first encountered difference.

## Modifying an Excel sheet

Add styles, formula and content to cells programmatically via a simple DSL.

``` java
@Test
public void shouldReplaceCellsInComplicatedAlternateSyntaxExample() throws IOException {
public void shouldReplaceCellsInComplicatedAlternateSyntaxExample() throws IOException {
    HSSFWorkbook workbook = getWorkbook("shouldReplaceCellsInComplicatedExampleTemplate.xls");
    new PoiWorkbookMutator(workbook)
            .replaceCell(coordinate(C, 5), "Adding")
            .replaceCell(coordinate(D, 11), "a")
            .replaceCell(coordinate(G, 3), "cell")
            .replaceCell(coordinate(J, 10), "total")
            .replaceCell(coordinate(M, 15), 99.99d);

    assertThat(workbook, sameWorkbook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplateExpected.xls")));
}
```

A break in the matcher would show something like

``` java
java.lang.AssertionError:
Expected: equality of cell "G1"
     but: cell at "G1" contained <"Text"> expected <99.99D>
	at org.hamcrest.MatcherAssert.assertThat(MatcherAssert.java:20)
```


### Adding Styling

Define some styles to reuse

``` java
private final Border border = border(top(NONE), right(THIN_SOLID), bottom(THIN_SOLID), left(THIN_SOLID));
private final DataFormat numberFormat = dataFormat("#,##0.00");
```

Next create your `Cell` with style information.

``` java
Cell cell = new DoubleCell(99.99d, aStyle().with(border).with(numberFormat));
```

and add it to a `Row`.

``` java
HashMap<ColumnIndex, Cell> columns = new HashMap<ColumnIndex, Cell>() {{
    put(column(A), cell);
}};
Row row = new Row(columns);
```

add your row to the workbook.

``` java
new PoiWorkbookMutator(workbook).mutator.appendRowToFirstSheet(row);
```



## Using Matchers

To get more detailed output from mismatches, be sure to use `MatcherAssert.assertThat` from Hamcrest rather than the vanilla JUnit version (`org.junit.Assert.assertThat`). If you use the JUnit version, you'll get output similar to the following.

``` java
java.lang.AssertionError:
Expected: entire workbook to be equal
     got: <org.apache.poi.hssf.usermodel.HSSFWorkbook@6405ce40>
```

When, using `MatcherAssert`, you'd see something like this.

``` java
java.lang.AssertionError:
Expected: entire workbook to be equal
     but: cell at "C14" contained <"bananas"> expected <nothing>,
          cell at "C15" contained <"1,850,000 EUR"> expected <"1,850,000.00 EUR">,
          cell at "D16" contained <nothing> expected <"Tue Sep 04 06:30:00">
	at org.hamcrest.MatcherAssert.assertThat(MatcherAssert.java:20)
```

### Caveats

   - Currently, matching doesn't take into account styling. A cell is equal to another regardless of style. If one has a border for example, and the other doesn't but they have the same values, they are considered equal.


# Download

Available for manual download via my [Maven repository](http://robotooling.com/maven/) or add the repository to your 'pom.xml'.

``` xml
<repositories>
    <repository>
        <id>bad.robot</id>
        <name>bad.robot repository on robotooling</name>
        <url>https://robotooling.com/maven/</url>
        <snapshots>
            <enabled>true</enabled>
            <updatePolicy>always</updatePolicy>
        </snapshots>
    </repository>
</repositories>
```


Then add a dependency.

``` xml
<dependency>
    <groupId>bad.robot</groupId>
    <artifactId>simple-excel</artifactId>
    <version>1.2</version>
    <scope>compile</scope>
</dependency>
```

For more tools, see [robotooling.com](http://www.robotooling.com) and visit my [blog](http://baddotrobot.com).

# Releasing

Take a look at the [release readme](RELEASE.md).
