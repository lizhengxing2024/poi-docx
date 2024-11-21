package com.dw.vsd2png;

import org.apache.poi.common.usermodel.PictureType;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.ooxml.POIXMLRelation;
import org.apache.poi.ooxml.util.DocumentHelper;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtListItem;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.springframework.boot.test.context.SpringBootTest;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

import java.io.*;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;

import static org.apache.poi.ooxml.POIXMLTypeLoader.DEFAULT_XML_OPTIONS;

@SpringBootTest
class DocxWriter {

    private void addVisio(XWPFDocument document) throws OpenXML4JException, NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException, IOException, SAXException, XmlException {
        /**
         * XWPFRelation
         *      type:
         *          application/vnd.ms-visio.drawing
         *      rel:
         *          POIXMLDocument.PACK_OBJECT_REL_TYPE
         *          http://schemas.openxmlformats.org/officeDocument/2006/relationships/package
         *      defaultName:
         *          "/word/embeddings/Microsoft_Visio___#.vsdx"
         */


        /**
         * 写入缩略图
         */
        InputStream picIs = new FileInputStream("D://visio//vsdx.png");
        byte[] picData = IOUtils.toByteArrayWithMaxLength(picIs, XWPFPictureData.getMaxImageSize());
        int picIdx = document.getNextPicNameNumber(PictureType.PNG);
        POIXMLRelation picRelDesc = XWPFRelation.IMAGE_PNG;
        XWPFPictureData xwpfPicData = (XWPFPictureData) document.createRelationship(picRelDesc, XWPFFactory.getInstance(), picIdx);
        PackagePart picDataPart = xwpfPicData.getPackagePart();
        try (OutputStream out = picDataPart.getOutputStream()) {
            out.write(picData);
        } catch (IOException e) {
            throw new POIXMLException(e);
        }
        String picRelationId = document.getRelationId(xwpfPicData);
        System.out.println("PIC:" + picRelationId);

        /**
         * 写入Visio
         */
        // 制造 Vsio-XWPFRelation
        POIXMLRelation.NoArgConstructor noArgConstructor = XWPFVisioData::new;
        POIXMLRelation.ParentPartConstructor parentPartConstructor = (parent, part) -> new XWPFVisioData();
        Class<?> clazz = XWPFRelation.class;
        Constructor<?> constructor = clazz.getDeclaredConstructor(String.class, String.class, String.class,
                POIXMLRelation.NoArgConstructor.class,
                POIXMLRelation.ParentPartConstructor.class);
        constructor.setAccessible(true);
        XWPFRelation visioRelDesc = (XWPFRelation) constructor.newInstance("application/vnd.ms-visio.drawing",
                POIXMLDocument.PACK_OBJECT_REL_TYPE,
                "/word/embeddings/Microsoft_Visio___#.vsdx",
                noArgConstructor, parentPartConstructor);
        // 写入
        InputStream visioIs = new FileInputStream("D://visio//vsdx.vsdx");
        byte[] visioData = IOUtils.toByteArrayWithMaxLength(visioIs, XWPFPictureData.getMaxImageSize());
        int visioIdx = document.getAllEmbeddedParts().size() + 1; // proposal: 根据现存嵌入的数量，生成的建议文件名，例如："/word/embeddings/Microsoft_Visio___1.vsdx"
        XWPFVisioData dwpfVisioData = (XWPFVisioData) document.createRelationship(visioRelDesc, XWPFFactory.getInstance(), visioIdx);
        PackagePart visioDataPart = dwpfVisioData.getPackagePart();
        try (OutputStream out = visioDataPart.getOutputStream()) {
            out.write(visioData);
        } catch (IOException e) {
            throw new POIXMLException(e);
        }
        String visioRelationId = document.getRelationId(dwpfVisioData);
        System.out.println("VISIO:" + visioRelationId);
//        XWPFVisioData test = (XWPFVisioData) document.getRelationById(relationId);

        /**
         * 写入XML
         */
        String shapeId = "_x0000_i1025";
        String shapeWidth = "415.15pt";
        String shapeHeight = "52.85pt";
        String objectXml = "<w:object xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                "               w:dxaOrig=\"9706\" w:dyaOrig=\"1232\">" +
                "           <v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                "                    xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
                "                    xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "               id=\"" + shapeId + "\"\n" +
                "               style=\"width:"+shapeWidth+";height:"+shapeHeight+"\" o:ole=\"\">\n" +
                "               <v:imagedata r:id=\"" + picRelationId + "\" o:title=\"\" />\n" +
                "           </v:shape>\n" +
                "           <o:OLEObject xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
                "                        xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "               Type=\"Embed\" ProgID=\"Visio.Drawing.15\" ShapeID=\"" + shapeId + "\"\n" +
                "               DrawAspect=\"Content\" ObjectID=\"_1793607303\" r:id=\"" + visioRelationId + "\" />" +
                "</w:object>";
        InputSource objXmlIs = new InputSource(new StringReader(objectXml));
        org.w3c.dom.Document objectDoc = DocumentHelper.readDocument(objXmlIs);
        XmlToken xmlObject = XmlToken.Factory.parse(objectDoc.getDocumentElement(), DEFAULT_XML_OPTIONS);
        CTR ctr = document.createParagraph().createRun().getCTR();
        ctr.set(xmlObject);
    }


    @Test
    public void writeTable() throws IOException, OpenXML4JException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException, XmlException, SAXException {
        XWPFDocument document = new XWPFDocument();

        // 插入visio
        this.addVisio(document);

        int rowCount = 6;
        int columnCount = 4;
        XWPFTable table = document.createTable(rowCount, columnCount);
        table.setWidth(11907 - 1800 - 1800); // A4纸张的可用宽度
        for (int row = 0; row < rowCount; row++) {
            XWPFTableRow tableRow = table.getRow(row);


//            // 插入下拉框
//            CTSdtCell ctSdtCell = tableRow.getCtRow().addNewSdt();
//            CTSdtListItem ctSdtListItem = ctSdtCell.addNewSdtPr().addNewDropDownList().addNewListItem();
//            ctSdtListItem.setDisplayText("XYZ");
//            ctSdtListItem.setValue("XYZ");
//            ctSdtCell.addNewSdtContent().addNewTc().addNewP().addNewR().addNewT().setStringValue("XYZ");


            for (int col = 0; col < columnCount; col++) {
                XWPFTableCell tableCell = tableRow.getCell(col);


                // 控制样式
                XWPFRun run = tableCell.getParagraphs().get(0).createRun();
                run.setBold(true);
                run.setText("Cell " + row + "," + col);

                // 插入图片
                FileInputStream imageStream = new FileInputStream("D:\\visio\\insertPng2Docx\\page1-页-1.png");
                tableCell.addParagraph().createRun().addPicture(imageStream, XWPFDocument.PICTURE_TYPE_JPEG, "image.jpg", Units.toEMU(50), Units.toEMU(50));
                imageStream.close();

                // 插入下拉框
                CTSdtRun ctSdtRun = tableCell.addParagraph().getCTP().addNewSdt();
                CTSdtListItem ctSdtListItem1 = ctSdtRun.addNewSdtPr().addNewDropDownList().addNewListItem();
                ctSdtListItem1.setDisplayText("XYZ");
                ctSdtListItem1.setValue("XYZ");
                ctSdtRun.addNewSdtContent().addNewR().addNewT().setStringValue("XYZ");


                // 合并单元格
                if (row > 3) {
                    if (col == 1) {
                        tableCell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
                    } else {
                        tableCell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
                    }
                }
            }
        }

        FileOutputStream out = new FileOutputStream("table_example.docx");
        document.write(out);
        out.close();
        document.close();
    }

    private class QuestionItem {
        public String label;
        public String value;

        // layout
        int rowspan;
    }
}
