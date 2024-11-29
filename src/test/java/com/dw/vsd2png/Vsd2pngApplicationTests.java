package com.dw.vsd2png;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.xdgf.usermodel.XmlVisioDocument;
import org.apache.poi.xdgf.usermodel.shape.ShapeRenderer;
import org.apache.poi.xdgf.util.VsdxToPng;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.boot.test.context.SpringBootTest;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.io.*;
import java.util.List;

@SpringBootTest
class Vsd2pngApplicationTests {


    // DOCX->PDF
    @Test
    public void docx2pdf() {
        String source = "D:\\visio\\extractVsdFromWord\\docx.docx";
        String target = "D:\\visio\\extractVsdFromWord\\docx.pdf";

        long start = System.currentTimeMillis();
        ActiveXComponent app = null;
        Dispatch doc = null;
        try {
            File targetFile = new File(target);
            if (targetFile.exists()) {
                targetFile.delete();
            }

            ComThread.InitSTA();
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", false);
            Dispatch docs = app.getProperty("Documents").toDispatch();

            System.out.println("打开文档" + source);
            doc = Dispatch.call(docs, "Open", source, false, true).toDispatch();

            System.out.println("转换文档到PDF " + target);
            Dispatch.call(doc, "SaveAs", target, 17); // wordSaveAsPDF为特定值17

            long end = System.currentTimeMillis();
            System.out.println("转换完成用时：" + (end - start) + "ms.");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (doc != null) {
                Dispatch.call(doc, "Close", false);
            }

            if (app != null) {
                app.invoke("Quit", 0); // 不保存待定的更改
            }

            ComThread.Release();
        }
    }


    // VSD -> VSDX
    @Test
    void vsd2vsdx() throws IOException {
        ActiveXComponent app = new ActiveXComponent("Visio.Application");
        Dispatch docs = app.getProperty("Documents").toDispatch();
        Dispatch doc = Dispatch.call(docs, "Open", "D:\\visio\\vsdx2png\\vsd.vsd").toDispatch();
        Dispatch.call(doc, "SaveAs", "D:\\visio\\vsdx2png\\vsdx-from-vsd.vsdx");
        Dispatch.call(doc, "Close");
        app.invoke("Quit");
    }

    // VSDX -> PNG
    @Test
    void vsdx2png() throws IOException {
        ShapeRenderer renderer = new ShapeRenderer();
        XmlVisioDocument doc = new XmlVisioDocument(new FileInputStream("D:\\visio\\vsdx2png\\vsdx-from-vsd.vsdx"));
        VsdxToPng.renderToPng(doc, "D:\\visio\\vsdx2png", 100.0, renderer);
    }

    // 从 DOCX 中提取 VSD --- 结构丢失
    @Test
    void extractVisioFromWord_nostruct() throws IOException, OpenXML4JException {
        FileInputStream fis = new FileInputStream("D:\\visio\\extractVsdFromWord\\docx-compat.docx");
        XWPFDocument doc = new XWPFDocument(fis);
        List<PackagePart> allEmbeddedParts = doc.getAllEmbeddedParts();
        int cnt = 0;
        for (PackagePart pPart : allEmbeddedParts) {
            FileOutputStream fos = new FileOutputStream("D:\\visio\\extractVsdFromWord\\visio-from-word_" + cnt + ".vsd");
            copyStream(pPart.getInputStream(), fos);
            cnt++;
        }
    }

    // 从 DOCX 中提取表格完整数据
    // 1. 各单元格内容（普通的、下拉框）
    // 2. Visio文件以及其所在位置
    //
    @Test
    void extractVisioFromWord_withstruct() throws IOException, OpenXML4JException {
        FileInputStream fis = new FileInputStream("D:\\visio\\extractVsdFromWord\\docx.docx");
        XWPFDocument doc = new XWPFDocument(fis);

        System.out.println("==========================================================");

        // 解析出内嵌Visio，并获知：id、inputstream
        PackagePart packagePart = doc.getPackagePart();
        System.out.println("getPartName:" + packagePart.getPartName());
        PackageRelationshipCollection relationships = packagePart.getRelationshipsByType(POIXMLDocument.OLE_OBJECT_REL_TYPE);
        for (int i = 0; i < relationships.size(); i++) {
            PackageRelationship rel = relationships.getRelationship(i);
            System.out.println("getTargetURI:" + rel.getTargetURI());
            String id = rel.getId();
            InputStream inputStream = doc.getPackagePart().getRelatedPart(rel).getInputStream();
            FileOutputStream fos = new FileOutputStream("D:\\visio\\extractVsdFromWord\\" + id + ".vsd");
            copyStream(inputStream, fos);
            System.out.println(">>>>>>>>>This File Contain Visio:" + id);
        }
        System.out.println("==========================================================");

        // 遍历表格
        // 包含假设：
        // 1. 每行单元格应该是偶数（2、4），如果有sdt单元格，也应该出现在偶数列
        // 2. 当 getTableICells 和 getTableCells 返回的单元格数量不一致时，说明当前行中包含 sdt 单元格
        //      2.1. sdt 单元格中的是 dropdownlist 组件
        //      2.2. 该行中的非 sdt 单元格不会包含 visio 嵌入元素，只含有普通文本
        // 3. 非 sdt 单元格中如果有多行，可能包含 visio 嵌入元素
        //
        List<XWPFTable> tables = doc.getTables();
        for (XWPFTable table : tables) {
            List<XWPFTableRow> rows = table.getRows();

            for (int i = 0; i < rows.size(); i++) {
                XWPFTableRow row = rows.get(i);

                if (row.getTableICells().size() > row.getTableCells().size()) {
                    // 当前行包含sdt单元格
                    // sdt 单元格会被计数在 iCells 中，cells 中没有计数
                    List<ICell> tableICells = row.getTableICells();

                    int sdtCellIndex = -1; // sdtCell计数器，第几个sdtCell
                    for (int j = 0; j < tableICells.size(); j++) {
                        ICell icell = tableICells.get(j);
                        if (icell instanceof XWPFSDTCell) {
                            sdtCellIndex++;
                            // 包含 sdt 单元格行的， sdt 单元格，认为其中是下拉框。

                            // 但是从 tableICells 接口中无法直接解析其中的下拉框
                            // 只能从行的底层接口重新解析，此时依赖规则：单元格应该是偶数，sdt单元格应该出现在偶数列
                            // 所以...第2个icell..对应第一个sdtcell； 第四个icell...对应第二个sdtcell
                            CTSdtCell[] sdtCellArray = row.getCtRow().getSdtArray();
                            assert sdtCellArray.length >= sdtCellIndex + 1;

                            // 下拉框选项在sdtPr中
                            // 下拉框当前选中项在sdtcontent中
                            CTSdtCell sdtCell = sdtCellArray[sdtCellIndex];
                            CTSdtPr sdtCellPr = sdtCell.getSdtPr();
                            CTSdtDropDownList dropDownList = sdtCellPr.getDropDownList();
                            CTSdtListItem[] dropDownListItemArray = dropDownList.getListItemArray();
                            for (CTSdtListItem item : dropDownListItemArray) {
                                System.out.println("-----dropdownitem:" + item.getDisplayText() + ":" + item.getValue());
                            }
                            String dropDownListContent = sdtCell.getSdtContent().getTcArray(0).getPArray(0).getRArray(0).getTArray(0).getStringValue();
                            System.out.println(dropDownListContent);
                        } else if (icell instanceof XWPFTableCell cell) {
                            // 包含 sdt 单元格行的，非 sdt 单元格，认为其中只有普通文本。
                            System.out.println(cell.getText());
                        }
                    }
                } else {
                    // 当前行不包含sdt单元格

                    List<XWPFTableCell> tableCells = row.getTableCells();
                    for (int j = 0; j < tableCells.size(); j++) {
                        XWPFTableCell cell = tableCells.get(j);

                        // cell底下可能有段落，也可能有表格
                        int tableIndex = -1;
                        CTTc ctTc = cell.getCTTc();
                        NodeList cellChildrenNodes = ctTc.getDomNode().getChildNodes();
                        for (int o = 0; o < cellChildrenNodes.getLength(); o++) {
                            Node cellChild = cellChildrenNodes.item(o);
                            String cellChildNodeName = cellChild.getLocalName();
                            if ("tbl".equals(cellChildNodeName)) {
                                tableIndex = o - 1; // 排除掉 tcPr 子元素
                                break;
                            }
                        }

                        if (tableIndex > -1) {
                            // cell里面含有table
                            List<XWPFParagraph> paragraphs = cell.getParagraphs();
                            for (int k = 0; k < paragraphs.size(); k++) {
                                XWPFParagraph paragraph = paragraphs.get(k);
                                CTP ctp = paragraph.getCTP();

                                // 缺陷是不能处理多个表格...后面再议
                                if (k >= tableIndex) {
                                    // 这里应该是个table
                                    System.out.println("!!!!!!!!!!!!!!!!!!!!!!!!这里应该是个表格>>" + k);
                                    tableIndex = Integer.MAX_VALUE; // 只用一次


                                    // 同时也假设内部table就是普通文字
                                    XWPFTable xwpfTable = cell.getTables().get(0);
                                    for (XWPFTableRow xwpfTableRow : xwpfTable.getRows()) {
                                        System.out.print("ROW:-----------");
                                        for (XWPFTableCell tableCell : xwpfTableRow.getTableCells()) {
                                            System.out.print(tableCell.getText() + ", ");
                                        }
                                        System.out.println();
                                    }

                                    System.out.println("!!!!!!!!!!!!!!!!!!!!!!!!这里应该是个表格>>END");


                                }

                                // 这些段落里面如果有空的，可能就包含嵌入结构
                                if (ctp.toString().contains("<o:OLEObject")) {
                                    List<CTR> rList = ctp.getRList();
                                    if (rList.size() == 1) {
                                        CTR ctr = rList.get(0);
                                        CTObject[] objectArray = ctr.getObjectArray();
                                        if (objectArray.length > 0) {
                                            CTObject ctObject = objectArray[0];
                                            Node domNode = ctObject.getDomNode();
                                            NodeList childNodes = domNode.getChildNodes();
                                            for (int m = 0; m < childNodes.getLength(); m++) {
                                                Node item = childNodes.item(m);
                                                if ("shape".equals(item.getLocalName())) {
                                                    NamedNodeMap attributes = item.getAttributes();
                                                    Node objectID = attributes.getNamedItem("style");
                                                    System.out.println("HERE DISPLAY VISIO STYLE:" + objectID.getNodeValue());
                                                } else if ("OLEObject".equals(item.getLocalName())) {
                                                    NamedNodeMap attributes = item.getAttributes();
                                                    Node objectID = attributes.getNamedItem("r:id");
                                                    System.out.println("HERE DISPLAY VISIO:" + objectID.getNodeValue());
                                                }
                                            }
                                        }
                                    }
                                } else if (ctp.toString().contains("<w:drawing")) {
                                    List<XWPFPicture> embeddedPictures = paragraph.getRuns().get(0).getEmbeddedPictures();
                                    for (XWPFPicture pic : embeddedPictures) {
                                        System.out.println("!!!!!!!!!!!!这里包含图片:" + pic.getWidth() + ";" + pic.getPictureData().getData().length);
                                    }
                                } else {
                                    // 段落中包含文字
                                    System.out.println(paragraph.getText());
                                }
                            }
                        } else {
                            // cell里面不包含table
                            // 但是这里面可能有图片
                            if (cell.getParagraphs().size() > 1) {
                                // 不包含sdt单元格的行，如果有多个段落，有理由怀疑里面包含visio元素
                                // 此时需要用底层接口探测，找到对应的 tc 结构
                                CTTc tc = row.getCtRow().getTcArray(j);
                                if (tc.toString().indexOf("<o:OLEObject") > -1) {

                                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                                    for (XWPFParagraph paragraph : paragraphs) {
                                        // 这些段落里面如果有空的，可能就包含嵌入结构
                                        if ("".equals(paragraph.getText())) {
                                            CTP ctp = paragraph.getCTP();
                                            List<CTR> rList = ctp.getRList();
                                            if (rList.size() == 1) {
                                                CTR ctr = rList.get(0);
                                                CTObject[] objectArray = ctr.getObjectArray();
                                                if (objectArray.length > 0) {
                                                    CTObject ctObject = objectArray[0];
                                                    Node domNode = ctObject.getDomNode();
                                                    NodeList childNodes = domNode.getChildNodes();
                                                    for (int k = 0; k < childNodes.getLength(); k++) {
                                                        Node item = childNodes.item(k);
                                                        if ("shape".equals(item.getLocalName())) {
                                                            NamedNodeMap attributes = item.getAttributes();
                                                            Node objectID = attributes.getNamedItem("style");
                                                            System.out.println("HERE DISPLAY VISIO STYLE:" + objectID.getNodeValue());
                                                        } else if ("OLEObject".equals(item.getLocalName())) {
                                                            NamedNodeMap attributes = item.getAttributes();
                                                            Node objectID = attributes.getNamedItem("r:id");
                                                            System.out.println("HERE DISPLAY VISIO:" + objectID.getNodeValue());
                                                        }
                                                    }
                                                }
                                            }
                                        } else {
                                            // 段落中包含文字
                                            System.out.println(paragraph.getText());
                                        }
                                    }
                                } else {
                                    // 里面只有纯文字
                                    System.out.println(cell.getText());
                                }
                            } else {
                                // 不包含sdt单元格的行，如果只有一个段落，则认为里面只有普通文字，不存在嵌入的
                                System.out.println(cell.getText());
                            }
                        }


                    }
                }
            }
        }
    }

    private void copyStream(InputStream inputStream, OutputStream outputStream) throws IOException {
        byte[] buffer = new byte[1024];
        int bytesRead;
        while ((bytesRead = inputStream.read(buffer)) != -1) {
            outputStream.write(buffer, 0, bytesRead);
        }
        outputStream.flush();
    }
}
