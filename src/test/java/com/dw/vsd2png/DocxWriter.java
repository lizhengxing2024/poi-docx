package com.dw.vsd2png;

import org.apache.poi.xwpf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

@SpringBootTest
class DocxWriter {
    @Test
    public void writeTable() throws IOException {
        // 创建一个新的空白DOCX文档
        XWPFDocument document = new XWPFDocument();

        // 创建一个表格
        XWPFTable table = document.createTable(2, 2); // 2行2列的表格


        // 设置表格的宽度
        table.setWidth(11907 - 1800 - 1800);

        // 添加表格数据
        for (int row = 0; row < 2; row++) {
            XWPFTableRow tableRow = table.getRow(row);
            for (int col = 0; col < 2; col++) {
                XWPFTableCell tableCell = tableRow.getCell(col);
                tableCell.setText("Cell " + row + "," + col);
            }
        }

        // 将文档写入到文件系统
        FileOutputStream out = new FileOutputStream("table_example.docx");
        document.write(out);
        out.close();

        // 关闭文档
        document.close();
    }
}
