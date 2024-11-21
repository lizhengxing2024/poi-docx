package com.dw.vsd2png;

import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.ooxml.POIXMLRelation;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtListItem;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.*;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;

@SpringBootTest
class DocxWriter {

    private void addVisio(XWPFDocument document) throws OpenXML4JException, NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException, IOException {
        InputStream is = new FileInputStream("D://visio//vsdx.vsdx");
        byte[] visioData = IOUtils.toByteArrayWithMaxLength(is, XWPFPictureData.getMaxImageSize());

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
         * proposal: 根据现存嵌入的数量，生成的建议文件名，例如："/word/embeddings/Microsoft_Visio___1.vsdx"
         */
        int idx = document.getAllEmbeddedParts().size() + 1;

        // 制造 XWPFRelation
        POIXMLRelation.NoArgConstructor noArgConstructor = XWPFVisioData::new;
        POIXMLRelation.ParentPartConstructor parentPartConstructor = (parent, part) -> new XWPFVisioData();
        Class<?> clazz = XWPFRelation.class;
        Constructor<?> constructor = clazz.getDeclaredConstructor(String.class, String.class, String.class,
                POIXMLRelation.NoArgConstructor.class,
                POIXMLRelation.ParentPartConstructor.class);
        constructor.setAccessible(true);
        XWPFRelation relDesc = (XWPFRelation) constructor.newInstance("application/vnd.ms-visio.drawing",
                POIXMLDocument.PACK_OBJECT_REL_TYPE,
                "/word/embeddings/Microsoft_Visio___#.vsdx",
                noArgConstructor, parentPartConstructor);

        // 创建XWPFVisioData
        XWPFVisioData dwpfVisioData = (XWPFVisioData) document.createRelationship(relDesc, XWPFFactory.getInstance(), idx);
        PackagePart picDataPart = dwpfVisioData.getPackagePart();
        try (OutputStream out = picDataPart.getOutputStream()) {
            out.write(visioData);
        } catch (IOException e) {
            throw new POIXMLException(e);
        }
        String relationId = document.getRelationId(dwpfVisioData);
//        XWPFVisioData test = (XWPFVisioData) document.getRelationById(relationId);
//        System.out.println(test);
//        document.createRelationship()

    }


    @Test
    public void writeTable() throws IOException, OpenXML4JException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
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
