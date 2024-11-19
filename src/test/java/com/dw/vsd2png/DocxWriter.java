package com.dw.vsd2png;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileOutputStream;
import java.io.IOException;

@SpringBootTest
class DocxWriter {
    @Test
    public void writeTable() throws IOException {
        XWPFDocument document = new XWPFDocument();

        int rowCount = 6;
        int columnCount = 4;
        XWPFTable table = document.createTable(rowCount, columnCount);
        table.setWidth(11907 - 1800 - 1800); // A4纸张的可用宽度
        for (int row = 0; row < rowCount; row++) {
            XWPFTableRow tableRow = table.getRow(row);
            for (int col = 0; col < columnCount; col++) {
                XWPFTableCell tableCell = tableRow.getCell(col);
                tableCell.setText("Cell " + row + "," + col);
            }
        }

        FileOutputStream out = new FileOutputStream("table_example.docx");
        document.write(out);
        out.close();
        document.close();
    }
}
