package com.dw.vsd2png;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import org.docx4j.Docx4J;
import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

@SpringBootTest

public class Docx2Html {

    @Test
    public void poi2html() throws Docx4JException, FileNotFoundException {
        long millis1 = System.currentTimeMillis();

        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File("docx.docx"));

        // 设置 HTML 转换选项
        HTMLSettings htmlSettings = Docx4J.createHTMLSettings();
        htmlSettings.setWmlPackage(wordMLPackage);
        htmlSettings.setImageDirPath("images/");
        htmlSettings.setImageTargetUri("images/");

        // 输出 HTML 文件
        FileOutputStream os = new FileOutputStream("docx4j.html");
        Docx4J.toHTML(htmlSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);

        //do something
        long millis2 = System.currentTimeMillis();
        long time = millis2 - millis1;//经过的毫秒数
        System.out.println(time / 1000);
    }

    @Test
    public void com2html() {
        String source = "D:\\visio\\extractVsdFromWord\\docx.docx";
        String target = "D:\\visio\\extractVsdFromWord\\htmlfiltered";

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
            // https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat
            Dispatch.call(doc, "SaveAs", target, 10); // wordSaveAsPDF为特定值17

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
}
