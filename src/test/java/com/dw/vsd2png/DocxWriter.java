package com.dw.vsd2png;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlObject;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

@SpringBootTest
class DocxWriter {
    @Test
    public void writeTable() throws IOException, InvalidFormatException {
        XWPFDocument document = new XWPFDocument();

        int rowCount = 6;
        int columnCount = 4;
        XWPFTable table = document.createTable(rowCount, columnCount);
        table.setWidth(11907 - 1800 - 1800); // A4纸张的可用宽度
        for (int row = 0; row < rowCount; row++) {
            XWPFTableRow tableRow = table.getRow(row);


            // 插入下拉框
            CTSdtCell ctSdtCell = tableRow.getCtRow().addNewSdt();
            CTSdtListItem ctSdtListItem = ctSdtCell.addNewSdtPr().addNewDropDownList().addNewListItem();
            ctSdtListItem.setDisplayText("XYZ");
            ctSdtListItem.setValue("XYZ");
            ctSdtCell.addNewSdtContent().addNewTc().addNewP().addNewR().addNewT().setStringValue("XYZ");


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
