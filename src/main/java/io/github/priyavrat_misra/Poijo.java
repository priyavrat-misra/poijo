package io.github.priyavrat_misra;

import io.github.priyavrat_misra.annotations.Column;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * A Java library featuring a fluent, annotation-driven API for working with Excel spreadsheets,
 * built atop Apache POI. Providing a cleaner approach to map, configure, style, and generate
 * workbooks with minimal boilerplate.
 *
 * @author Priyavrat Misra
 */
public class Poijo {
  public static final String SPACE = " ";
  public static final String EMPTY = "";

  private static final Logger logger = LoggerFactory.getLogger(Poijo.class);

  private final Workbook workbook;

  private Poijo(Workbook workbook) {
    this.workbook = workbook;
  }

  /**
   * The main entry point for using {@link Poijo} methods.
   *
   * @param workbook a {@link Workbook} on which operations are to be performed
   * @return a newly created {@link Poijo} object for chaining
   */
  public static Poijo using(Workbook workbook) {
    if (workbook == null) {
      logger.error("workbook is null");
      throw new NullPointerException("workbook cannot be null");
    }
    return new Poijo(workbook);
  }

  /**
   * Maps {@code object} to {@link Workbook} that was passed in {@link Poijo#using(Workbook)}.
   *
   * <p>Structurally the {@code object} is considered a workbook, and the {@link
   * java.util.Collection} data members are considered as sheets, where each element is a row. Refer
   * example below for more clarity.
   *
   * <p>Supports the following types:
   *
   * <ul>
   *   <li>types that are supported by {@link org.apache.poi.ss.usermodel.Cell}{@code ::setCell}
   *       ({@link String}, {@link Integer}, {@code int}, {@link Double}, {@code double}, {@link
   *       Boolean}, {@code boolean}, {@link RichTextString}, {@link Date}, {@link LocalDate},
   *       {@link LocalDateTime} or {@link Calendar})
   *   <li>{@link java.util.Collection}
   *   <li>nested POJOs (should be annotated with {@link Column#nested()} set to {@code true})
   * </ul>
   *
   * <p>Nested objects and lists are handled in a recursive manner, resulting in a flattened
   * representation, suitable for the two-dimensional view of a spreadsheet. The title for a nested
   * field is the path of field names to it from the base class delimited by {@link
   * io.github.priyavrat_misra.annotations.Workbook#delimiter()}.
   *
   * <p>Note: Only {@code public} fields are considered because in order to access other kind of
   * variables, the accessibility level has to be altered via reflection, and altering or bypassing
   * the accessibility of classes, methods, or fields through reflection violates the encapsulation
   * principle.
   *
   * <p>Example:
   *
   * <p>Say there are two classes {@code Author} and {@code Book}, and a base class, {@code Library}
   * (should be annotated with {@link io.github.priyavrat_misra.annotations.Workbook}).
   *
   * <pre><code>
   * {@literal @}Workbook // to indicate this class will be used as a Workbook
   * public class Library {
   *     // custom sheet name can be provided using {@literal @}Sheet(name = "...")
   *     public Set&lt;Book&gt; books;
   * }
   * </code></pre>
   *
   * <pre><code>
   * {@literal @}Order({"title", "author", "publicationDate", "price"}) // column ordering
   * public class Book {
   *     public String title;
   *
   *     {@literal @}Column(nested = true) // to indicate nesting
   *     public Author author;
   *
   *     {@literal @}Column(numberFormat = "[$$-409]#,##0;-[$$-409]#,##0") // custom number format
   *     public double price;
   *
   *     {@literal @}Column(name = "Date of Publication", numberFormat = "dd/MM/yyyy") // custom title
   *     public LocalDate publicationDate;
   *
   *     public Book(String title, Author author, double price, LocalDate publicationDate) {
   *         this.title = title;
   *         this.author = author;
   *         this.price = price;
   *         this.publicationDate = publicationDate;
   *     }
   * }
   * </code></pre>
   *
   * <pre>{@code
   * public class Author {
   *     public String name;
   *     public List<String> genres;
   *
   *     public Author(String name, List<String> genres) {
   *         this.name = name;
   *         this.genres = genres;
   *     }
   * }
   * }</pre>
   *
   * <p>Now all it takes is something as simple as the following:
   *
   * <pre>{@code
   * public class Main {
   *     public static void main(String[] args) {
   *         // prepare data
   *         Library library = new Library();
   *         library.books = Arrays.asList(
   *           new Book(
   *             "The Hobbit",
   *             new Author("J.R.R. Tolkien", Arrays.asList("Fantasy", "Adventure")),
   *             14.99,
   *             LocalDate.parse("1937-09-21")),
   *           new Book(
   *             "Harry Potter and the Sorcerer's Stone",
   *             new Author("J.K. Rowling", Arrays.asList("Fantasy", "Drama", "Young Adult")),
   *             19.99,
   *             LocalDate.parse("1997-06-26")));
   *
   *         try (Workbook workbook = new XSSFWorkbook();
   *             OutputStream fileOut = Files.newOutputStream(Paths.get("workbook.xlsx"))) {
   *           // this is where the magic happens!
   *           Poijo.using(workbook)
   *              .map(library)
   *              .applyCellStyleProperties(Collections.singletonMap(CellPropertyType.ALIGNMENT, HorizontalAlignment.CENTER))
   *              .write(fileOut);
   *         } catch (IOException e) {
   *           throw new RuntimeException(e);
   *         }
   *     }
   * }
   * }</pre>
   *
   * <p>The spreadsheet will look something like the following in a sheet named <i>"Books"</i>.
   *
   * <table style='text-align: center;'>
   *   <tr>
   *     <th>Title</th>
   *     <th>Author Name</th>
   *     <th>Author Genres 0</th>
   *     <th>Author Genres 1</th>
   *     <th>Author Genres 2</th>
   *     <th>Date of Publication</th>
   *     <th>Price</th>
   *   </tr>
   *   <tr>
   *     <td>The Hobbit</td>
   *     <td>J.R.R. Tolkien</td>
   *     <td>Fantasy</td>
   *     <td>Adventure</td>
   *     <td></td>
   *     <td>21/09/1937</td>
   *     <td>$15</td>
   *   </tr>
   *   <tr>
   *     <td>Harry Potter and the Sorcerer's Stone</td>
   *     <td>J.K. Rowling</td>
   *     <td>Fantasy</td>
   *     <td>Drama</td>
   *     <td>Young Adult</td>
   *     <td>26/06/1997</td>
   *     <td>$20</td>
   *   </tr>
   * </table>
   *
   * <p>For header and body specific styling, {@link #applyCellStylePropertiesToHeader} and {@link
   * #applyCellStylePropertiesToBody} can be used respectively.
   *
   * @param object an object which is to be mapped to the {@link Workbook} passed in {@link
   *     Poijo#using(Workbook)}
   * @return {@code this} instance for chaining
   * @param <T> type parameter for {@code object}
   * @see io.github.priyavrat_misra.annotations.Workbook
   * @see io.github.priyavrat_misra.annotations.Sheet
   * @see io.github.priyavrat_misra.annotations.Column
   * @see io.github.priyavrat_misra.annotations.Order
   * @see Poijo#applyCellStylePropertiesToHeader(Map)
   * @see Poijo#applyCellStylePropertiesToBody(Map)
   * @see <a
   *     href="https://wiki.sei.cmu.edu/confluence/display/java/SEC05-J.+Do+not+use+reflection+to+increase+accessibility+of+classes%2C+methods%2C+or+fields"
   *     target="_blank">SEC05-J. Do not use reflection to increase accessibility of classes,
   *     methods, or fields</a>
   */
  public <T> Poijo map(T object) {
    PojoMapper.map(workbook, validate(object));
    return this;
  }

  /**
   * @throws NullPointerException if {@code object} is {@code null}
   * @throws IllegalArgumentException if {@link T} is not annotated with {@link
   *     io.github.priyavrat_misra.annotations.Workbook}
   */
  private <T> T validate(T object) {
    if (object == null) {
      logger.error("object is null");
      throw new NullPointerException("object cannot be null");
    }
    if (!object
        .getClass()
        .isAnnotationPresent(io.github.priyavrat_misra.annotations.Workbook.class)) {
      logger.error(
          "{} is not annotated with io.github.priyavrat_misra.annotations.Workbook",
          object.getClass().getName());
      throw new IllegalArgumentException(
          "Passed object's class is not annotated with io.github.priyavrat_misra.annotations.Workbook");
    }
    return object;
  }

  /**
   * Invokes {@link CellUtil#setCellStylePropertiesEnum(Cell, Map)} on each cell in the sheets of
   * the workbook.
   *
   * @param styles the properties to be added to a cell style, as {{@link CellPropertyType}:
   *     propertyValue}
   * @return {@code this} instance for chaining
   * @see CellUtil#setCellStylePropertiesEnum(Cell, Map)
   */
  public Poijo applyCellStyleProperties(Map<CellPropertyType, Object> styles) {
    workbook.forEach(
        sheet ->
            sheet.forEach(
                row -> row.forEach(cell -> CellUtil.setCellStylePropertiesEnum(cell, styles))));
    return this;
  }

  /**
   * Invokes {@link CellUtil#setCellStylePropertiesEnum(Cell, Map)} on each cell of the first row in
   * the sheets of the workbook.
   *
   * @param styles the properties to be added to a cell style, as {{@link CellPropertyType}:
   *     propertyValue}
   * @return {@code this} instance for chaining
   * @see CellUtil#setCellStylePropertiesEnum(Cell, Map)
   */
  public Poijo applyCellStylePropertiesToHeader(Map<CellPropertyType, Object> styles) {
    workbook.forEach(
        sheet ->
            sheet.getRow(0).forEach(cell -> CellUtil.setCellStylePropertiesEnum(cell, styles)));
    return this;
  }

  /**
   * Invokes {@link CellUtil#setCellStylePropertiesEnum(Cell, Map)} on each cell except the first
   * row in the sheets of the workbook.
   *
   * @param styles the properties to be added to a cell style, as {{@link CellPropertyType}:
   *     propertyValue}
   * @return {@code this} instance for chaining
   * @see CellUtil#setCellStylePropertiesEnum(Cell, Map)
   */
  public Poijo applyCellStylePropertiesToBody(Map<CellPropertyType, Object> styles) {
    workbook.forEach(
        sheet -> {
          Iterator<Row> rowIterator = sheet.rowIterator();
          rowIterator.next(); // skip header
          rowIterator.forEachRemaining(
              row -> row.forEach(cell -> CellUtil.setCellStylePropertiesEnum(cell, styles)));
        });
    return this;
  }

  /**
   * Write out {@link Workbook} passed in {@link Poijo#using(Workbook)} to an {@link OutputStream}.
   *
   * @param stream the java OutputStream you wish to write to
   * @throws IOException if anything can't be written
   */
  public void write(OutputStream stream) throws IOException {
    workbook.write(stream);
  }

  /**
   * Get the passed workbook instance.
   *
   * @return the {@link Workbook} instance passed in {@link Poijo#using(Workbook)}
   */
  public Workbook getWorkbook() {
    return workbook;
  }
}
