# Poijo - A POJO to Excel Mapper Library

## Overview

**Poijo** is a Java library that simplifies the process of mapping nested Plain Old Java Objects (POJOs) to Excel spreadsheets using [Apache POI](https://poi.apache.org/). The library's name is derived from the combination of **POI** + **JO** = **Poijo**.

If you're wondering about Apache POI's name, it has a rather *punny* origin. POI stands for **Poor Obfuscation Implementation**, a tongue-in-cheek reference to Microsoft's binary file formats for Office documents. So, if you're ever frustrated with Excel's inner workings, rest assured that even the creators of POI had a laugh about it. And now, with **Poijo**, you can have some fun too—by making Excel work for you instead of the other way around!

### Features

- **Annotation-based configuration**: Use annotations to define workbook, sheet, and column properties.
- **Support for multiple sheets**: Map collections of objects into separate sheets.
- **Nested objects or collections support**: Automatically flatten and map nested objects or collections into spreadsheet columns.
- **Custom sheet and column ordering**: Specify the order of sheets and columns using `@Order`.
- **Custom sheet and column names**: Define custom names for sheets and columns or don't, both are handled.
- **Number formatting**: Apply custom number formats to columns (e.g., dates, currencies, phone numbers, zip codes, etc.).
- **Type-safe mapping**: Only supports types compatible with Apache POI's `Cell#setCellValue`.

### Limitations

- **Public Members Only**: Poijo only works with `public` fields. While non-public fields could be accessed via reflection by modifying their accessibility, this approach has been avoided to respect the encapsulation principle.

## Usage

### Example

1. Define POJO classes

> ```java
> @Workbook
> public class DTO {
>   @Sheet(name = "Books")
>   public List<Book> books;
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

2. Map POJOs to a workbook

> ```java
> import org.apache.poi.ss.usermodel.Workbook;
> import org.apache.poi.xssf.usermodel.XSSFWorkbook;
> import io.github.priyavrat_misra.Poijo;
> 
> import java.io.OutputStream;
> import java.nio.file.Files;
> import java.nio.file.Paths;
> import java.time.LocalDate;
> import java.util.Arrays;
> 
> public class Main {
>   public static void main(String[] args) {
>     DTO dto = new DTO();
>     dto.books =
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
>                 19.99));
> 
>     try (Workbook workbook = new XSSFWorkbook();
>         OutputStream fileOut = Files.newOutputStream(Paths.get("workbook.xlsx"))) {
>       Poijo.map(workbook, dto).write(fileOut); // the magic happens here!
>     } catch (Exception e) {
>       throw new RuntimeException(e);
>     }
>   }
> }
> ```

3. The generated Excel file will contain a sheet named `Books` with the following structure:

> | Title                                 | Author Name    | Author Genres 0 | Author Genres 1 | Author Genres 2 | Date of Publication | Price |
> |---------------------------------------|----------------|-----------------|-----------------|-----------------|---------------------|-------|
> | The Hobbit                            | J.R.R. Tolkien | Fantasy         | Adventure       |                 | 21/09/1937          | $15   |
> | Harry Potter and the Sorcerer's Stone | J.K. Rowling   | Fantasy         | Drama           | Young Adult     | 26/06/1997          | $20   |

## Logging

This library uses [SLF4J](https://www.slf4j.org/) as a logging façade. It does not include a specific logging backend, allowing users to configure their preferred logging implementation (e.g., Logback, Log4j, etc.).

### How to Configure Logging

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

3. Run your application. Logs generated by **Poijo** will be handled by the logging backend you configured.

### Logging Levels

- **DEBUG**: Detailed information, including eligible fields, column names, and sheet creation.
- **INFO**: General information about the mapping process.
- **WARN**: Warnings about missing data or ordering issues in sheets/columns.
- **ERROR**: Errors such as null input or missing annotations.

## When to use Poijo?

If you're tired of manually creating cells, managing indexes, and debugging Excel files that look like they were formatted by a cat walking on your keyboard, **Poijo** is here to save the day. With just a few annotations, you can turn your POJOs into beautifully formatted Excel spreadsheets that even your boss will be impressed with.

## When not to use Poijo?

While Poijo simplifies the mapping of POJOs to Excel, it does so using Java's reflection, which can incur a runtime performance penalty. If you're working on mission-critical applications (e.g., designed for space missions), Poijo may not be suitable (though, frankly, you probably shouldn't be using Java or Excel for that either).

That being said, if you're already using a framework like Spring that relies on reflection, adding Poijo to your project likely won't introduce significant overhead, and the improved maintainability it offers may outweigh the performance tradeoff.

Ultimately, the decision is yours.

## Contributing

Contributions are welcome! Please fork the repository, create a feature branch, and submit a pull request. Bonus points if you add a new pun to the README.

## License

This project is licensed under the [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0.txt).