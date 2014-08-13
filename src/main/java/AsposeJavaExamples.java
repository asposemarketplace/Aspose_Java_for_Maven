import com.aspose.words.*;
import com.aspose.words.Font;
import com.aspose.words.Shape;

import java.awt.*;

/**
 * Created by Adeel Ilyas on 6/28/2014.
 */
public class AsposeJavaExamples {

    public static void runAsposeWord(String wordOutPutFile) throws Exception    {
        com.aspose.words.Document doc = new com.aspose.words.Document();

        // DocumentBuilder provides members to easily add content to a
        // document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertImage(AsposeAPITest.class.getResourceAsStream("images/aspose.png"));

        shape.setWrapType(WrapType.TOP_BOTTOM);


        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);

        shape.setHorizontalAlignment(HorizontalAlignment.CENTER);


        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PARAGRAPH);

        shape.setVerticalAlignment(VerticalAlignment.TOP);


        Font font = builder.getFont();

        font.setSize(16);

        font.setColor(java.awt.Color.BLUE);

        font.setName("Arial");


        builder.insertParagraph();
        builder.insertParagraph();
        // Save the document in DOCX format. The format to save as is
        // inferred from the extension of the file name.
        // Aspose.Words supports saving any document in many more formats.

        builder.startTable();
        builder.insertCell();

        // Set height and define the height rule for the header row.
        builder.getRowFormat().setHeight(40.0);
        builder.getRowFormat().setHeightRule(HeightRule.AT_LEAST);

        // Some special features for the header row.
        builder.getCellFormat()
                .getShading()
                .setBackgroundPatternColor(
                        new java.awt.Color(198, 217, 241));
        builder.getParagraphFormat()
                .setAlignment(ParagraphAlignment.CENTER);
        builder.getFont().setSize(16);
        builder.getFont().setName("Arial");
        builder.getFont().setBold(true);

        builder.getCellFormat().setWidth(100.0);
        builder.write("First Name");
        builder.insertCell();
        builder.write("Last Name");
        builder.insertCell();
        builder.write("Address");
        builder.insertCell();
        builder.write("City");
        builder.insertCell();
        builder.write("Telephone");
        builder.endRow();
        // Set features for the other rows and cells.
        builder.getCellFormat().getShading()
                .setBackgroundPatternColor(java.awt.Color.WHITE);
        builder.getCellFormat().setWidth(100.0);
        builder.getCellFormat().setVerticalAlignment(
                CellVerticalAlignment.CENTER);

        // Reset height and define a different height rule for table body
        builder.getRowFormat().setHeight(30.0);
        builder.getRowFormat().setHeightRule(HeightRule.AUTO);

        builder.insertCell();
        // Reset font formatting.
        builder.getFont().setSize(12);
        builder.getFont().setBold(false);
        builder.write("Adeel");
        builder.insertCell();
        builder.write("Ilyas");
        builder.insertCell();
        builder.write("SC 97 Sector 31 F, Darussalam Society Korangi Crossing");
        builder.insertCell();
        builder.write("Karachi");
        builder.insertCell();
        builder.write("923222694809");
        builder.endRow();

        builder.endTable();
        builder.insertParagraph();
        builder.insertParagraph();
        builder.getFont().setSize(7);
        builder.getFont().setName("Arial");
        builder.getFont().setBold(false);
        builder.getFont().setColor(Color.black);
        builder.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<project xmlns=\"http://maven.apache.org/POM/4.0.0\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"\n" +
                "\txsi:schemaLocation=\"http://maven.apache.org/POM/4.0.0 http://maven.apache.org/xsd/maven-4.0.0.xsd\">\n" +
                "\t<modelVersion>4.0.0</modelVersion>\n" +
                "\t<groupId>com.aspose</groupId>\n" +
                "\t<artifactId>aspose.maven.example</artifactId>\n" +
                "\t<packaging>jar</packaging>\n" +
                "\t<version>1.0</version>\n" +
                "\t<name>Aspose for Maven Example</name>\n" +
                "\t<repositories>\n" +
                "\t\t<repository>\n" +
                "\t\t      <id>AsposeJavaAPI</id>\n" +
                "\t\t      <name>Aspose Java API</name>\n" +
                "\t\t      <url>http://maven.aspose.com/artifactory/simple/ext-release-local/</url>\n" +
                "\t\t</repository>\n" +
                "\t</repositories>\n" +
                "\t<dependencies>\n" +
                "\t\t<dependency>\n" +
                "\t\t      <groupId>com.aspose</groupId>\n" +
                "\t\t      <artifactId>aspose-words</artifactId>\n" +
                "\t\t      <version>14.5.0</version>\n" +
                "\t\t</dependency>\n" +
                "\t\t<dependency>\n" +
                "\t\t      <groupId>junit</groupId>\n" +
                "\t\t      <artifactId>junit</artifactId>\n" +
                "\t\t      <version>3.8.1</version>\n" +
                "\t\t      <scope>test</scope>\n" +
                "\t\t</dependency>\n" +
                "\t</dependencies>\n" +
                "</project>");
        // Save the document

        doc.save(wordOutPutFile, com.aspose.words.SaveFormat.DOC);
    }
}
