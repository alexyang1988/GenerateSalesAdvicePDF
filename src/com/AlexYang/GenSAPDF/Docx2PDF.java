package com.AlexYang.GenSAPDF;

import java.io.File;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;


public class Docx2PDF
{
    static final int wdDoNotSaveChanges = 0;// 不保存待定的更改。
    static final int wdFormatPDF = 17;// PDF 格式

    public static void docx2PDF()
    {
        String filename = new File("doc/output.docx").getAbsolutePath();
        String tofilename = new File("doc/output.pdf").getAbsolutePath();

        ActiveXComponent app = null;
        try {
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", false);

            Dispatch docs = app.getProperty("Documents").toDispatch();
            System.out.println("打开文档..." + filename);
            Dispatch doc = Dispatch.call(docs,//
                    "Open",//
                    filename,// FileName
                    false,// ConfirmConversions
                    true // ReadOnly
            ).toDispatch();

            System.out.println("转换文档到PDF..." + tofilename);
            File tofile = new File(tofilename);
            if (tofile.exists())
            {
                tofile.delete();
            }
            Dispatch.call(doc,//
                    "SaveAs", //
                    tofilename, // FileName
                    wdFormatPDF);

            Dispatch.call(doc, "Close", false);

        } 
        catch (Exception e) 
        {
            System.out.println("Error:文档转换失败：" + e.getMessage());
        } 
        finally 
        {
            if (app != null)
                app.invoke("Quit", wdDoNotSaveChanges);
        }
    }
}
