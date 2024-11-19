package com.example.dynamic.map.controller;

import com.aspose.pdf.DocSaveOptions;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.w3c.dom.Document;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.*;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/account-details/dynamic")
public class PdfToDocxController {

    @GetMapping("generateDocxFromDynamicXml")
    public ResponseEntity<byte[]> generateDocxFromDynamicXml() throws Exception {
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
    }

    @GetMapping("/generateDocxFromXml")
    public ResponseEntity<?> generateDocxFromXml() throws Exception {
        // Load and parse the XML template
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document document = builder.parse(new File("D:\\MY PROJECT\\dynamic-map\\src\\main\\resources\\templates\\input.xml"));

        // Create a new DOCX document
        XWPFDocument docxDocument = new XWPFDocument();

        // Process the entire XML document
        processNode(document.getDocumentElement(), docxDocument);

        Map<String,Object> res = new HashMap<>();
        // Save the DOCX document to a byte array
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            docxDocument.write(out);
            return ResponseEntity.ok()
                    .header("Content-Disposition", "attachment; filename=output.docx")
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
                    .body(out.toByteArray());
        }
    }

    private void processNode(Node node, XWPFDocument docxDocument) {
        String nodeName = node.getNodeName();

        if (nodeName.equals("fo:block")) {
            processBlock(node, docxDocument);
        } else {
            processChildren(node, docxDocument);
        }
    }

    private void processChildren(Node node, XWPFDocument docxDocument) {
        NodeList children = node.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            processNode(children.item(i), docxDocument);
        }
    }

    private void processBlock(Node blockNode, XWPFDocument docxDocument) {
        XWPFParagraph paragraph = docxDocument.createParagraph();
        // Create a run to hold the text styling
        XWPFRun run = paragraph.createRun();

        // Apply alignment based on attributes
        NamedNodeMap attributes = blockNode.getAttributes();
        if (attributes != null) {
            Node textAlign = attributes.getNamedItem("text-align");
            if (textAlign != null) {
                String alignment = textAlign.getNodeValue();
                if ("center".equalsIgnoreCase(alignment)) paragraph.setAlignment(ParagraphAlignment.CENTER);
                else if ("right".equalsIgnoreCase(alignment)) paragraph.setAlignment(ParagraphAlignment.RIGHT);
                else paragraph.setAlignment(ParagraphAlignment.LEFT);
            }

            Node underLine = attributes.getNamedItem("text-decoration");
            if (underLine != null) {
                run.setUnderline(UnderlinePatterns.SINGLE);
            }

            // Set color
            Node colorNode = attributes.getNamedItem("color");
            if (colorNode != null) {
                String colorValue = colorNode.getNodeValue().replace("#","");
                run.setColor(colorValue); // Use hex color code, e.g., "FF0000" for red
            }

            // Set font-family
            Node fontFamilyNode = attributes.getNamedItem("font-family");
            if (fontFamilyNode != null) {
                String fontFamily = fontFamilyNode.getNodeValue();
                run.setFontFamily(fontFamily);
            }

            // Set font-size
            Node fontSizeNode = attributes.getNamedItem("font-size");
            if (fontSizeNode != null) {
                try {
                    int fontSize = Integer.parseInt(fontSizeNode.getNodeValue().replace("pt", "").trim());
                    run.setFontSize(fontSize);
                } catch (NumberFormatException e) {
                    System.err.println("Invalid font size: " + fontSizeNode.getNodeValue());
                }
            }

            // Set font-weight (bold)
            Node fontWeightNode = attributes.getNamedItem("font-weight");
            if (fontWeightNode != null) {
                String fontWeight = fontWeightNode.getNodeValue();
                run.setBold("bold".equalsIgnoreCase(fontWeight));
            } else run.setBold(false);

            // Set margin-top (Spacing before paragraph)
            Node marginTopNode = attributes.getNamedItem("margin-top");
            if (marginTopNode != null) {
                try {
                    int marginTop = Integer.parseInt(marginTopNode.getNodeValue().replace("px", "").trim());
                    paragraph.setSpacingBefore(marginTop * 15); // Convert points to twips (1px = 15 twips)
                } catch (NumberFormatException e) {
                    System.err.println("Invalid margin-top: " + marginTopNode.getNodeValue());
                }
            }
        }

        // Add content to the paragraph or process nested tables
        NodeList children = blockNode.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node child = children.item(i);
            if ("fo:table".equals(child.getNodeName())) {
                processTable(child, docxDocument,blockNode);
            }
            /*else if ("fo:leader".equals(child.getNodeName())) {
                addLeader(docxDocument);
            }*/
            else if ("fo:inline".equals(child.getNodeName())) {
                processFoInline(child, paragraph,run);
            }
            else if ("fo:block".equals(child.getNodeName())) {
                processBlock(child,docxDocument);

            } else {
                String text = child.getTextContent().trim();
                if (!text.isEmpty()) run.setText(text);
            }
        }
    }

    private void processBlockInsideOfCell(Node blockNode, XWPFTableCell docxDocument, Node ownerBlock) {
        // Create a new paragraph inside the table cell
        XWPFParagraph paragraph = docxDocument.addParagraph();
        XWPFRun run = paragraph.createRun();

        // Apply styling from attributes of the owner block
        NamedNodeMap attributes = ownerBlock.getAttributes();
        if (attributes != null) {
            // Set text alignment
            Node textAlign = attributes.getNamedItem("text-align");
            if (textAlign != null) {
                switch (textAlign.getNodeValue().toLowerCase()) {
                    case "center" -> paragraph.setAlignment(ParagraphAlignment.CENTER);
                    case "justify" -> paragraph.setAlignment(ParagraphAlignment.BOTH);
                    case "right" -> paragraph.setAlignment(ParagraphAlignment.RIGHT);
                    default -> paragraph.setAlignment(ParagraphAlignment.LEFT);
                }
            }

            // Set underline style
            Node underline = attributes.getNamedItem("text-decoration");
            if (underline != null && "underline".equalsIgnoreCase(underline.getNodeValue())) {
                run.setUnderline(UnderlinePatterns.SINGLE);
            }

            // Set text color
            Node colorNode = attributes.getNamedItem("color");
            if (colorNode != null) {
                run.setColor(colorNode.getNodeValue().replace("#", "")); // Remove "#" if present
            }

            // Set font family
            Node fontFamilyNode = attributes.getNamedItem("font-family");
            if (fontFamilyNode != null) {
                run.setFontFamily(fontFamilyNode.getNodeValue());
            }

            // Set font size
            Node fontSizeNode = attributes.getNamedItem("font-size");
            if (fontSizeNode != null) {
                try {
                    int fontSize = Integer.parseInt(fontSizeNode.getNodeValue().replace("pt", "").trim());
                    run.setFontSize(fontSize);
                } catch (NumberFormatException e) {
                    System.err.println("Invalid font size: " + fontSizeNode.getNodeValue());
                }
            }

            // Set bold if font-weight is specified
            Node fontWeightNode = attributes.getNamedItem("font-weight");
            run.setBold(fontWeightNode != null && "bold".equalsIgnoreCase(fontWeightNode.getNodeValue()));

            // Set spacing before the paragraph (margin-top)
            Node marginTopNode = attributes.getNamedItem("margin-top");
            if (marginTopNode != null) {
                try {
                    int marginTop = Integer.parseInt(marginTopNode.getNodeValue().replace("px", "").trim());
                    paragraph.setSpacingBefore(marginTop * 15); // Convert px to twips (1px = 15 twips)
                } catch (NumberFormatException e) {
                    System.err.println("Invalid margin-top: " + marginTopNode.getNodeValue());
                }
            }
        }

        // Process child nodes of the block node
        NodeList children = blockNode.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node child = children.item(i);

            switch (child.getNodeName()) {
                case "fo:inline" -> processFoInline(child, paragraph, run);
                case "fo:list-block" -> processListBlock(child, docxDocument, ownerBlock);
                case "fo:block" -> processBlockInsideOfCell(child, docxDocument, ownerBlock);
                default -> {
                    String text = child.getTextContent().trim();
                    if (!text.isEmpty()) run.setText(text);
                }
            }
        }
    }

    private void processListBlock(Node listBlock, XWPFTableCell docxDocument, Node ownerBlock) {
        NodeList listItems = listBlock.getChildNodes();
        for (int i = 0; i < listItems.getLength(); i++) {
            Node listItem = listItems.item(i);
            if ("fo:list-item".equals(listItem.getNodeName())) {
                NodeList listItemChildren = listItem.getChildNodes();
                for (int j = 0; j < listItemChildren.getLength(); j++) {
                    Node child = listItemChildren.item(j);
                    if ("fo:list-item-label".equals(child.getNodeName()) || "fo:list-item-body".equals(child.getNodeName())) {
                        NodeList blocks = child.getChildNodes();
                        for (int k = 0; k < blocks.getLength(); k++) {
                            Node block = blocks.item(k);
                            if ("fo:block".equals(block.getNodeName())) {
                                processBlockInsideOfCell(block, docxDocument, ownerBlock);
                            }
                        }
                    }
                }
            }
        }
    }


    private void processFoInline(Node child, XWPFParagraph paragraph,XWPFRun run) {
        NamedNodeMap namedNodeMap = child.getAttributes();
        if (namedNodeMap != null) {
            Node fontWeightNode = namedNodeMap.getNamedItem("font-weight");
            if (fontWeightNode != null) {
                String fontWeight = fontWeightNode.getNodeValue();
                run.setBold("bold".equalsIgnoreCase(fontWeight));
            } else {
                run.setBold(false);
            }

            // Set margin-left (Spacing before paragraph)
            Node marginLeftNode = namedNodeMap.getNamedItem("margin-left");
            if (marginLeftNode != null) {
                try {
                    int marginTop = Integer.parseInt(marginLeftNode.getNodeValue().replace("pt", "").trim());
                    paragraph.setSpacingBefore(marginTop * 20); // Convert points to twips (1pt = 20 twips)
                } catch (NumberFormatException e) {
                    System.err.println("Invalid margin-top: " + marginLeftNode.getNodeValue());
                }
            }
        }
        String text = child.getTextContent().trim();
        if (!text.isEmpty()) run.setText(text);
    }

    private void processTable(Node tableNode, XWPFDocument docxDocument,Node blockNode) {
        XWPFTable table = docxDocument.createTable();
        /*NamedNodeMap attribute = tableNode.getAttributes();
        if (attribute != null) {

            Node widthNode = attribute.getNamedItem("width");
            String width = widthNode.getNodeValue();
            // Set table width
            table.setWidth("100%"); // Set table width to 100%

            // Set borders for the table
            CTTbl ctTable = table.getCTTbl();
            CTTblBorders borders = ctTable.addNewTblPr().addNewTblBorders();
            borders.addNewTop().setVal(STBorder.SINGLE);
            borders.addNewBottom().setVal(STBorder.SINGLE);
            borders.addNewLeft().setVal(STBorder.SINGLE);
            borders.addNewRight().setVal(STBorder.SINGLE);
            borders.addNewInsideH().setVal(STBorder.SINGLE);
            borders.addNewInsideV().setVal(STBorder.SINGLE);
        }*/
        NodeList tableChildren = tableNode.getChildNodes();

        // Process table columns for widths
        Map<Integer, String> columnWidths = processTableColumns(tableNode);

        for (int i = 0; i < tableChildren.getLength(); i++) {
            Node child = tableChildren.item(i);
            if ("fo:table-body".equals(child.getNodeName())) {
                createTableBody(child, table, columnWidths,blockNode,docxDocument);
            }
        }
    }

    private Map<Integer, String> processTableColumns(Node tableNode) {
        Map<Integer, String> columnWidths = new HashMap<>();
        NodeList children = tableNode.getChildNodes();

        int columnIndex = 0;
        for (int i = 0; i < children.getLength(); i++) {
            Node child = children.item(i);
            if ("fo:table-column".equals(child.getNodeName())) {
                NamedNodeMap attributes = child.getAttributes();
                Node widthAttr = attributes.getNamedItem("column-width");
                if (widthAttr != null) {
                    columnWidths.put(columnIndex++, widthAttr.getNodeValue());
                }
            }
        }
        return columnWidths;
    }

    private void createTableBody(Node tableBodyNode, XWPFTable table, Map<Integer, String> columnWidths,Node blockNode,XWPFDocument document) {
        NodeList rowList = tableBodyNode.getChildNodes();
        int rowIndex = 0;
        for (int i = 0; i < rowList.getLength(); i++) {
            Node rowNode = rowList.item(i);
            if ("fo:table-row".equals(rowNode.getNodeName())) {
                XWPFTableRow tableRow = (rowIndex >= table.getNumberOfRows()) ? table.createRow() : table.getRow(rowIndex);
                createTableRow(rowNode, tableRow, columnWidths,blockNode,document);
            }
            rowIndex++;
        }
    }

    private void createTableRow(Node rowNode, XWPFTableRow tableRow, Map<Integer, String> columnWidths,Node blockNode,XWPFDocument document) {
        NodeList cellList = rowNode.getChildNodes();
        int cellIndex = 0;

        for (int i = 0; i < cellList.getLength(); i++) {
            Node cellNode = cellList.item(i);
            if ("fo:table-cell".equals(cellNode.getNodeName())) {
                XWPFTableCell tableCell = (cellIndex >= tableRow.getTableCells().size())
                        ? tableRow.addNewTableCell()
                        : tableRow.getCell(cellIndex);

                NodeList cellChild = cellNode.getChildNodes();
                for (int j = 0 ; j < cellChild.getLength() ; j ++) {
                    Node childNode = cellChild.item(j);
                    if ("fo:block".equals(childNode.getNodeName())) {
                        processBlockInsideOfCell(childNode,tableCell,blockNode);
                    }
                }
                /*String text = cellNode.getTextContent().trim();
                XWPFParagraph paragraph = tableCell.addParagraph();
                // Create a run to hold the text styling
                XWPFRun run = paragraph.createRun();

                // Apply alignment based on attributes
                NamedNodeMap attributes = blockNode.getAttributes();
                if (attributes != null) {
                    Node textAlign = attributes.getNamedItem("text-align");
                    if (textAlign != null) {
                        String alignment = textAlign.getNodeValue();
                        if ("center".equalsIgnoreCase(alignment)) paragraph.setAlignment(ParagraphAlignment.CENTER);
                        else if ("right".equalsIgnoreCase(alignment)) paragraph.setAlignment(ParagraphAlignment.RIGHT);
*//*
                        else if ("justify".equalsIgnoreCase(alignment)) paragraph.setAlignment(ParagraphAlignment.BOTH);
*//*
                        else paragraph.setAlignment(ParagraphAlignment.LEFT);
                    }

                    // Apply padding-like effects based on attributes
                    *//*Node paddingNode = attributes.getNamedItem("padding");
                    if (paddingNode != null) {
                        String paddingValue = paddingNode.getNodeValue();
                        int padding = Integer.parseInt(paddingValue.replace("pt", "").trim()); // Assuming padding is in "pt"

                        // Set spacing before and after (simulating padding)
                        paragraph.setSpacingBefore(padding);
                        paragraph.setSpacingAfter(padding);
                    }

                    Node paddingLeftNode = attributes.getNamedItem("padding-left");
                    if (paddingLeftNode != null) {
                        String paddingLeftValue = paddingLeftNode.getNodeValue();
                        int paddingLeft = Integer.parseInt(paddingLeftValue.replace("pt", "").trim()); // Assuming padding is in "pt"

                        // Set left indentation (simulating padding)
                        paragraph.setIndentationLeft(paddingLeft);
                    }*//*


                    // Set color
                    Node colorNode = attributes.getNamedItem("color");
                    if (colorNode != null) {
                        String colorValue = colorNode.getNodeValue().replace("#","");
                        run.setColor(colorValue); // Use hex color code, e.g., "FF0000" for red
                    }

                    *//*Node underLine = attributes.getNamedItem("text-decoration");
                    if (underLine != null) {
                        run.setUnderline(UnderlinePatterns.SINGLE);
                    }*//*

                    // Set font-family
                    Node fontFamilyNode = attributes.getNamedItem("font-family");
                    if (fontFamilyNode != null) {
                        String fontFamily = fontFamilyNode.getNodeValue();
                        run.setFontFamily(fontFamily);
                    }

                    // Set font-size
                    Node fontSizeNode = attributes.getNamedItem("font-size");
                    if (fontSizeNode != null) {
                        try {
                            int fontSize = Integer.parseInt(fontSizeNode.getNodeValue().replace("pt", "").trim());
                            run.setFontSize(fontSize);
                        } catch (NumberFormatException e) {
                            System.err.println("Invalid font size: " + fontSizeNode.getNodeValue());
                        }
                    }

                    // Set font-weight (bold)
                    Node fontWeightNode = attributes.getNamedItem("font-weight");
                    if (fontWeightNode != null) {
                        String fontWeight = fontWeightNode.getNodeValue();
                        run.setBold("bold".equalsIgnoreCase(fontWeight));
                    } else run.setBold(false);

                    // Set margin-top (Spacing before paragraph)
                    Node marginTopNode = attributes.getNamedItem("margin-top");
                    if (marginTopNode != null) {
                        try {
                            int marginTop = Integer.parseInt(marginTopNode.getNodeValue().replace("px", "").trim());
                            paragraph.setSpacingBefore(marginTop * 15); // Convert points to twips (1pt = 20 twips)
                        } catch (NumberFormatException e) {
                            System.err.println("Invalid margin-top: " + marginTopNode.getNodeValue());
                        }
                    }
                }


                if (!text.isEmpty()) run.setText(text);*/

                if (columnWidths.containsKey(cellIndex)) {
                    applyColumnWidth(tableCell, columnWidths.get(cellIndex));
                }
                cellIndex++;
            }
        }
    }

    private void applyColumnWidth(XWPFTableCell cell, String width) {
        if (width != null) {
            try {
                int widthInTwips = 0;

                // Handle different width formats
                if (width.endsWith("in")) {
                    // Convert inches to twips
                    double inches = Double.parseDouble(width.replace("in", "").trim());
                    widthInTwips = (int) (inches * 72 * 20);
                } else if (width.endsWith("pt")) {
                    // Convert points to twips
                    double points = Double.parseDouble(width.replace("pt", "").trim());
                    widthInTwips = (int) (points * 20);
                } else if (width.endsWith("%")) {
                    // Handle percentage widths (Assume a default table width, e.g., 5000 twips)
                    double percentage = Double.parseDouble(width.replace("%", "").trim());
                    int tableWidthInTwips = 5000; // Default table width in twips
                    widthInTwips = (int) ((percentage / 100) * tableWidthInTwips);
                } else {
                    throw new IllegalArgumentException("Unsupported width format: " + width);
                }

                // Apply the calculated width
                cell.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(widthInTwips));
            } catch (IllegalArgumentException e) {
                System.err.println("Invalid column width: " + width);
            }
        }
    }


    /*private void addLeader(XWPFDocument docxDocument) {
        XWPFParagraph paragraph = docxDocument.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.addBreak();
    }*/



    @GetMapping("/pdfToDocxFromPyScript")
    public ResponseEntity<byte[]> pdfToDocxFromPyScript(@RequestParam MultipartFile file) {
        try {
            // Save uploaded PDF to a temporary file
            File pdfFile = File.createTempFile("input", ".pdf");
            try (FileOutputStream out = new FileOutputStream(pdfFile)) {
                out.write(file.getBytes());
            }

            // Set output DOCX file
            File docxFile = File.createTempFile("output", ".docx");

            // Run Python script to convert PDF to DOCX
            ProcessBuilder processBuilder = new ProcessBuilder(
                    "python3", "pdf_to_docx.py",
                    pdfFile.getAbsolutePath(),
                    docxFile.getAbsolutePath()
            );
            processBuilder.redirectErrorStream(true);
            Process process = processBuilder.start();
            process.waitFor();

            // Read the converted DOCX file into a byte array
            byte[] docxBytes = java.nio.file.Files.readAllBytes(docxFile.toPath());

            // Clean up temporary files
            pdfFile.delete();
            docxFile.delete();

            // Return the DOCX file as a response
            return ResponseEntity.ok()
                    .header("Content-Disposition", "attachment; filename=output.docx")
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
                    .body(docxBytes);
        } catch (IOException | InterruptedException e) {
            e.printStackTrace(); // Log the exception
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("Error converting PDF to DOCX".getBytes());
        }
    }

    @GetMapping("convertPDFToDOCXFromPdBox")
    public ResponseEntity<byte[]> convertPDFToDOCXFromPdBox(@RequestParam MultipartFile file) {
        try (InputStream pdfBytes = file.getInputStream()) {
            // Convert PDF byte[] to DOCX byte[]
            byte[] docxBytes = convertPDFToDOCX(pdfBytes);

            // Return the DOCX file as a byte array
            return ResponseEntity.ok()
                    .header("Content-Disposition", "attachment; filename=output.docx")
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
                    .body(docxBytes);
        } catch (IOException e) {
            // Log the exception (consider using a logger)
            e.printStackTrace(); // Replace with proper logging
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("Error converting PDF to DOCX".getBytes());
        } catch (NullPointerException e) {
            // Handle case where the resource is not found
            return ResponseEntity.status(HttpStatus.NOT_FOUND)
                    .body("PDF file not found".getBytes());
        }
    }

    public byte[] convertPDFToDOCX(InputStream pdfBytes) throws IOException {
        // Step 1: Extract text from PDF
        String extractedText = extractTextFromPDF(pdfBytes);

        // Step 2: Convert extracted text to DOCX
        return createDOCXFromText(extractedText);
    }

    public byte[] createDOCXFromText(String extractedText) throws IOException {
        XWPFDocument document = new XWPFDocument();

        // Example logic to determine if extracted text contains table data
        String[] lines = extractedText.split("\n");
        boolean isTable = false;
        List<String[]> tableData = new ArrayList<>();

        for (String line : lines) {
            if (line.startsWith("|")) { // Assuming table rows start with '|'
                isTable = true;
                String[] rowData = line.split("\\|"); // Split by pipe character
                tableData.add(rowData);
            } else {
                if (isTable) {
                    // Create a table in DOCX
                    createTableInDOCX(document, tableData);
                    isTable = false; // Reset for next potential table
                    tableData.clear(); // Clear previous data
                }
                // Add paragraph for non-table text
                XWPFParagraph paragraph = document.createParagraph();
                paragraph.createRun().setText(line);
            }
        }

        // Handle any remaining table data
        if (!tableData.isEmpty()) {
            createTableInDOCX(document, tableData);
        }

        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            document.write(out);
            return out.toByteArray();
        }
    }

    private void createTableInDOCX(XWPFDocument document, List<String[]> tableData) {
        XWPFTable table = document.createTable(tableData.size(), tableData.get(0).length);

        for (int i = 0; i < tableData.size(); i++) {
            String[] rowData = tableData.get(i);
            for (int j = 0; j < rowData.length; j++) {
                XWPFTableCell cell = table.getRow(i).getCell(j);
                cell.setText(rowData[j].trim());
            }
        }
    }

    public String extractTextFromPDF(InputStream pdfBytes) throws IOException {
        try (PDDocument document = PDDocument.load(pdfBytes)) {
            PDFTextStripper stripper = new PDFTextStripper();
            return stripper.getText(document); // Extracts text content
        }
    }

    @GetMapping("pdfToDocx")
    public ResponseEntity<byte[]> convertPDFToDOCX(@RequestParam MultipartFile file) throws IOException {
        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(new ByteArrayInputStream(file.getBytes()));

        // Set save options for DOCX format
        DocSaveOptions saveOptions = new DocSaveOptions();
        saveOptions.setFormat(DocSaveOptions.DocFormat.DocX);
        saveOptions.setMode(DocSaveOptions.RecognitionMode.Flow); // Ensures flowable structure

        // Create output stream to hold DOCX data
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        // Save the document to the output stream in DOCX format
        pdfDocument.save(outputStream, saveOptions);

        // Return DOCX as response
        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=output.docx")
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
                .body(outputStream.toByteArray());

    }
}
