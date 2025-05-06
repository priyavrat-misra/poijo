# Poijo

## Overview

**Poijo** is a Java library with a [fluent interface](https://en.wikipedia.org/wiki/Fluent_interface) for working with Excel workbooks, built on top of [Apache POI](https://poi.apache.org/). It offers an annotation-driven API for mapping POJOs to Excel, as well as fluent utilities for styling and manipulating spreadsheets. _Poijo_ aims to simplify both exporting data and working with existing Excel files, and is designed for extensibility—so stay tuned for more features!

### What's in a Name?

_Poijo_ is a portmanteau of Apache _"POI"_ and _"POJO"_ (Plain Old Java Object).

Curious about Apache POI’s own name? *"POI"* stands for _Poor Obfuscation Implementation_—a playful jab at how tricky Microsoft’s Office file formats can be. So, if you ever find Excel confusing, you’re in good company!

### Features

- **Annotation-based configuration**: Use annotations to define workbook, sheet, and column properties.
- **Spreadsheet styling**: Apply cell styles, formats, and other modifications to any workbook, not just those created via mapping.
- **Multiple sheets mapping**: Map collections into separate sheets.
- **Nested object and collection support**: Automatically flatten and map nested objects or collections into spreadsheet columns.
- **Custom sheet and column ordering**: Specify the order of sheets and columns.
- **Custom sheet and column names**: Define custom names for sheets and columns—or don’t! Both are handled gracefully.
- **Number formatting**: Apply custom number formats to columns (dates, currencies, phone numbers, zip codes, and more).
- **Type-safe mapping**: Supports only types compatible with Apache POI's `Cell#setCellValue`.
- **Extensible design**: The foundation is laid for even more Excel-related features in the future.

### Limitations

- **Public Members Only**: _Poijo_ only works with `public` fields. While non-public fields could be accessed via reflection by modifying their accessibility, this approach has been avoided to respect the encapsulation principle.

---

## Requirements

- Java 8+
- [Apache POI](https://poi.apache.org/) (included as a dependency)

---

## Usage

Below are some quick-start examples. For more detailed documentation, see the Javadoc.

### Example: Mapping Data

1. Prepare POJOs

> ```java
> @Workbook // optional if there are no arguments
> public class Library {
>   @Sheet(name = "Books")
>   public List<Book> books;
>
>   public Library(List<Book> books) {
>       this.books = books;
>   }
> }
> 
> @Order({"title", "author", "publicationDate", "price"}) // column ordering
> public class Book {
>   public String title;
> 
>   @Column(nested = true) // to indicate nesting
>   public Author author;
> 
>   @Column(numberFormat = "[$$-409]#,##0;-[$$-409]#,##0") // custom number format
>   public double price;
> 
>   @Column(name = "Date of Publication", numberFormat = "dd/MM/yyyy") // custom title
>   public LocalDate publicationDate;
> 
>   public Book(String title, Author author, double price, LocalDate publicationDate) {
>     this.title = title;
>     this.author = author;
>     this.price = price;
>     this.publicationDate = publicationDate;
>   }
> }
> 
> public class Author {
>   public String name;
>   public List<String> genres;
> 
>   public Author(String name, List<String> genres) {
>     this.name = name;
>     this.genres = genres;
>   }
> }
> ```
>
> With this, your POJOs are ready for spreadsheet mapping. No need to manually create cells, worry about off-by-one errors, or deal with the usual spreadsheet headaches.

2. Map POJOs to a workbook and apply styles

> ```java
> public class Main {
>   public static void main(String[] args) {
> 
>     // prepare data
>     Library library = new Library(
>         Arrays.asList(
>             new Book(
>                 "The Hobbit",
>                 new Author("J.R.R. Tolkien", Arrays.asList("Fantasy", "Adventure")),
>                 LocalDate.parse("1937-09-21"),
>                 14.99),
>             new Book(
>                 "Harry Potter",
>                 new Author("J.K. Rowling", Arrays.asList("Fantasy", "Drama", "Young Adult")),
>                 LocalDate.parse("1997-06-26"),
>                 19.99)));
>
>     // map data
>     try (Workbook workbook = new XSSFWorkbook();
>         OutputStream fileOut = Files.newOutputStream(Paths.get("workbook.xlsx"))) {
>       Poijo.using(workbook)
>         .map(library)
>         .applyCellStyleProperties(Map.of(CellPropertyType.ALIGNMENT, HorizontalAlignment.CENTER))
>         .write(out);
>     } catch (Exception e) {
>       throw new RuntimeException(e);
>     }
>   }
> }
> ```

3. The generated Excel file will contain a sheet named `Books` with the following structure:

> |                 Title                 |  Author Name   | Author Genres 0 | Author Genres 1 | Author Genres 2 | Date of Publication | Price |
> |:-------------------------------------:|:--------------:|:---------------:|:---------------:|:---------------:|:-------------------:|:-----:|
> |              The Hobbit               | J.R.R. Tolkien |     Fantasy     |    Adventure    |                 |     21/09/1937      |  $15  |
> | Harry Potter and the Sorcerer's Stone |  J.K. Rowling  |     Fantasy     |      Drama      |   Young Adult   |     26/06/1997      |  $20  |

### Example: Styling an Existing Workbook

```java
public class Main {
  public static void main(String[] args) {
    try (Workbook workbook = WorkbookFactory.create(new File("existing.xlsx"));
        OutputStream styled = Files.newOutputStream(Paths.get("styled.xlsx"))) {
      Poijo.using(workbook)
          .applyCellStyleProperties(
              Map.of(
                  CellPropertyType.ALIGNMENT, HorizontalAlignment.CENTER,
                  CellPropertyType.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND))
          .applyCellStylePropertiesToHeader(
              Map.of(CellPropertyType.FILL_FOREGROUND_COLOR, IndexedColors.LIME.getIndex()))
          .applyCellStylePropertiesToBody(
              Map.of(CellPropertyType.FILL_FOREGROUND_COLOR, IndexedColors.WHITE.getIndex()))
          .write(styled);
    }
  }
}

```

---

## Logging

This library uses [SLF4J](https://www.slf4j.org/) as a logging façade. It does not include a specific logging backend, allowing you to choose your favorite logging implementation (Logback, Log4j, or whatever else sparks joy).

### Configuring Logging

1. Add SLF4J and your preferred logging backend to your project's dependencies. For example, to use [Logback](https://logback.qos.ch/), include the following in your `pom.xml`.

> ```xml
> 
> <dependencies>
>   <!-- SLF4J API -->
>   <dependency>
>     <groupId>org.slf4j</groupId>
>     <artifactId>slf4j-api</artifactId>
>     <version>2.0.17</version>
>   </dependency>
>   
>   <!-- Logback Classic -->
>   <dependency>
>     <groupId>ch.qos.logback</groupId>
>     <artifactId>logback-classic</artifactId>
>     <version>1.5.18</version>
>   </dependency>
> </dependencies>
> ```

2. Create a configuration file for your logging backend. For Logback, create a `logback.xml` file in the `src/main/resources` directory. Example:

> ```xml
> 
> <configuration>
>   <appender name="CONSOLE" class="ch.qos.logback.core.ConsoleAppender">
>     <encoder>
>       <pattern>%d{yyyy-MM-dd HH:mm:ss} [%thread] %-5level %logger{36} - %msg%n</pattern>
>     </encoder>
>   </appender>
> 
>   <root level="info">
>     <appender-ref ref="CONSOLE"/>
>   </root>
> 
>   <logger name="io.github.priyavrat_misra" level="debug"/> <!-- adjust as needed -->
> </configuration>
> ```

3. Run your application. Logs generated by _Poijo_ will be handled by your configured backend.

### Logging Levels

- **DEBUG**: Detailed information, including eligible fields, column names, and sheet creation.
- **INFO**: General information about the mapping process.
- **WARN**: Warnings about missing data or issues in sheet/column ordering.
- **ERROR**: Errors such as null input or missing annotations.

---

## Contributing

Contributions are welcome! Please follow these guidelines:

- _Poijo_ uses [google-java-format](https://github.com/google/google-java-format); please format your code with it before submitting PRs.
- Provide JUnit tests for your changes, preferably using [AssertJ](https://assertj.github.io/doc/)—let’s keep things fluent.
- Make sure your changes don’t break existing tests by running `mvn test`.
- Don’t rewrite history. (Version control is not a time machine. Yet.)

---

## License

This project is licensed under the [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0.txt).
