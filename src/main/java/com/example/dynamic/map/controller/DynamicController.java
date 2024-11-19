package com.example.dynamic.map.controller;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

@RestController
@RequestMapping("/api/account-details/dynamic")
public class DynamicController {

    /*public ResponseEntity<byte[]> generateDocxFromDynamicXml() throws Exception {
        InputStream xmlInput = getClass().getResourceAsStream("/templates/input.xml");
        // Parse XML
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        org.w3c.dom.Document doc = builder.parse(xmlInput);
        doc.getDocumentElement().normalize();

        // Create a new DOCX document
        try (XWPFDocument document = new XWPFDocument()) {
            // Create a table in the document
            XWPFTable table = document.createTable();

            // Extract the root element and its child nodes
            var root = doc.getDocumentElement();
            var childNodes = root.getChildNodes();

            // Create table header based on the first record's child nodes
            if (childNodes.getLength() > 0) {
                XWPFTableRow headerRow = table.getRow(0);
                var firstRecord = childNodes.item(0);
                var firstRecordChildren = firstRecord.getChildNodes();

                for (int i = 0; i < firstRecordChildren.getLength(); i++) {
                    if (firstRecordChildren.item(i).getNodeType() == org.w3c.dom.Node.ELEMENT_NODE) {
                        headerRow.addNewTableCell().setText(firstRecordChildren.item(i).getNodeName());
                    }
                }
            }

            // Populate the table with XML data
            for (int i = 0; i < childNodes.getLength(); i++) {
                var record = childNodes.item(i);
                if (record.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE) {
                    XWPFTableRow row = table.createRow();
                    var recordChildren = record.getChildNodes();

                    for (int j = 0; j < recordChildren.getLength(); j++) {
                        if (recordChildren.item(j).getNodeType() == org.w3c.dom.Node.ELEMENT_NODE) {
                            row.getCell(j).setText(recordChildren.item(j).getTextContent());
                        }
                    }
                }
            }

            // Write the document to a byte array
            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                document.write(outputStream);
                return ResponseEntity.ok().body(outputStream.toByteArray());
            }
        }
    }*/

    @GetMapping("/getFile")
    public byte[] generateDocxFromXml() throws Exception {
        // Load and parse the XML template
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document document = builder.parse(new File("D:\\MY PROJECT\\dynamic-map\\src\\main\\resources\\templates\\input.xml"));

        // Create a new DOCX document
        XWPFDocument docxDocument = new XWPFDocument();

// Add a title
        docxDocument.createParagraph().createRun().setText("TERMSHEET");

// Extract content
        NodeList blockList = document.getElementsByTagName("fo:block");
        for (int i = 0; i < blockList.getLength(); i++) {
            String paragraphText = blockList.item(i).getTextContent();
            docxDocument.createParagraph().createRun().setText(paragraphText);
        }

// Extract table
        NodeList tableList = document.getElementsByTagNameNS("http://www.w3.org/1999/XSL/Format", "table");
        if (tableList.getLength() > 0) {
            XWPFTable table = docxDocument.createTable();

            NodeList rowList = tableList.item(0).getChildNodes();
            for (int i = 0; i < rowList.getLength(); i++) {
                if (rowList.item(i).getNodeName().equals("fo:table-row")) {
                    XWPFTableRow tableRow;
                    if (i == 0) {
                        tableRow = table.getRow(0); // Get the first row for the header
                    } else {
                        tableRow = table.createRow(); // Create a new row for subsequent rows
                    }

                    NodeList cellList = rowList.item(i).getChildNodes();
                    for (int j = 0; j < cellList.getLength(); j++) {
                        if (cellList.item(j).getNodeName().equals("fo:table-cell")) {
                            String cellText = cellList.item(j).getTextContent();
                            XWPFTableCell tableCell = tableRow.getCell(j);
                            tableCell.setText(cellText);

                            // Optional: Style the header row differently
                            if (i == 0) {
                                XWPFParagraph cellParagraph = tableCell.addParagraph();
                                XWPFRun cellRun = cellParagraph.createRun();
                                cellRun.setText(cellText);
                                cellRun.setBold(true);
                                cellRun.setFontSize(12);
                                cellParagraph.setAlignment(ParagraphAlignment.CENTER);
                            }
                        }
                    }
                }
            }
        }

// Save the DOCX document to a byte array
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            docxDocument.write(out);
            return out.toByteArray();
        }
        }

    private void addContentFromXml(org.w3c.dom.Document doc, XWPFDocument document) {
        // Extract and add content from the XML document to the DOCX
        // This will dynamically read the XML structure and add it to the document.

        // Add the main content
        addMainContent(doc, document);

        // Add tables and other structured content
        addTablesFromXml(doc, document);
    }

    private void addMainContent(org.w3c.dom.Document doc, XWPFDocument document) {
        // Extracting the main content from the XML
        String mainContent = doc.getElementsByTagName("fo:block").item(0).getTextContent();
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(mainContent);
        run.addBreak();
    }

    private void addTablesFromXml(org.w3c.dom.Document doc, XWPFDocument document) {
        // Extracting tables from the XML
        var tables = doc.getElementsByTagName("fo:table");
        for (int i = 0; i < tables.getLength(); i++) {
            var tableElement = tables.item(i);
            XWPFTable table = document.createTable();

            // Extract rows from the table
            var rows = tableElement.getChildNodes();
            for (int j = 0; j < rows.getLength(); j++) {
                if (rows.item(j).getNodeName().equals("fo:table-row")) {
                    XWPFTableRow tableRow = table.createRow();
                    var cells = rows.item(j).getChildNodes();
                    for (int k = 0; k < cells.getLength(); k++) {
                        if (cells.item(k).getNodeName().equals("fo:table-cell")) {
                            String cellText = cells.item(k).getTextContent().trim();
                            if (j == 0 && k == 0) {
                                // Set header cell
                                tableRow.getCell(k ).setText(cellText);
                            } else {
                                // Set regular cell
                                tableRow.getCell(k).setText(cellText);
                            }
                        }
                    }
                }
            }
        }
    }

}
