---
title: Java中docx、xlsx、pptx转pdf的方法
tags: 
    - Java
    - Jacob
    - LibreOffice
    - pdf
categories: Java
date: 2023-01-08 13:13:58
---



## 一、Windows环境调用Office或者WPS的COM接口

### 引入jacob依赖
1. 下载[jacob-1.20依赖包](https://github.com/freemansoft/jacob-project/releases/tag/Root_B-1_20)并解压
2. 将其中的jar包安装到maven本地仓库（也可以直接在项目中引入jar包依赖）
    ```shell
    mvn install:install-file 
        -Dfile=</your/path/to/jacob.jar>
        -DgroupId=com.jacob
        -DartifactId=jacob
        -Dversion=1.20 -Dpackaging=jar
    ```
3. 把jacob-1.20-x64.dll和jacob-1.20-x86.dll拷贝至`%JAVA_HOME%\jre\bin`目录

> 关于jacob的更多使用方式，可以参考[MicroSoft的VBA文档](https://learn.microsoft.com/zh-cn/office/vba/api/word.documents.open)

### 转换代码：
```java
public class JacobToPdfUtil {
    /**
     * 转PDF格式值
     */
    private static final int WORD_FORMAT_PDF = 17;
    private static final int EXCEL_FORMAT_PDF = 0;
    private static final int PPT_FORMAT_PDF = 32;

    private static final String PROGRAM_WORD_OFFICE = "Word.Application";
    private static final String PROGRAM_WORD_WPS = "KWPS.Application";
    private static final String PROGRAM_EXCEL_OFFICE = "Excel.Application";
    private static final String PROGRAM_EXCEL_WPS = "KET.Application";
    private static final String PROGRAM_PPT_OFFICE = "PowerPoint.Application";
    private static final String PROGRAM_PPT_WPS = "KWPP.Application";

    public static void wordToPdfByMsOffice(String inputFile, String pdfFile) throws IOException {
        wordToPdf(inputFile, pdfFile, PROGRAM_WORD_OFFICE);
    }

    public static void wordToPdfByWps(String inputFile, String pdfFile) throws IOException {
        wordToPdf(inputFile, pdfFile, PROGRAM_WORD_WPS);
    }

    private static void wordToPdf(String inputFile, String pdfFile, String program) throws IOException {
        ActiveXComponent app = null;
        Dispatch doc = null;
        try {
            // 创建一个word对象
            app = new ActiveXComponent(program);
            // 不可见打开word
            app.setProperty("Visible", new Variant(false));
            // 禁用宏
            app.setProperty("AutomationSecurity", new Variant(3));
            // 获取文挡属性
            Dispatch docs = app.getProperty("Documents").toDispatch();
            // 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
            doc = Dispatch.call(docs, "Open", inputFile).toDispatch();
            // word保存为pdf格式宏，值为17
            Dispatch.call(doc, "SaveAs", pdfFile, WORD_FORMAT_PDF);
        } catch (Throwable t) {
            throw new IOException("word转pdf失败", t);
        } finally {
            if (doc != null) {
                Dispatch.call(doc, "Close", false);
            }
            if (app != null) {
                app.invoke("Quit");
            }
            // 关闭WINWORD.exe进程
            ComThread.Release();
        }
    }

    public static void excelToPdfByMsOffice(String inputFile, String pdfFile) throws IOException {
        excelToPdf(inputFile, pdfFile, PROGRAM_EXCEL_OFFICE);
    }

    public static void excelToPdfByWps(String inputFile, String pdfFile) throws IOException {
        excelToPdf(inputFile, pdfFile, PROGRAM_EXCEL_WPS);
    }

    private static void excelToPdf(String inputFile, String pdfFile, String program) throws IOException {
        ActiveXComponent app = null;
        Dispatch excel = null;
        try {
            app = new ActiveXComponent(program);
            // 不可见打开excel
            app.setProperty("Visible", new Variant(false));
            // 禁用宏
            app.setProperty("AutomationSecurity", new Variant(3));
            Dispatch workbooks = app.getProperty("Workbooks").toDispatch();
            excel = Dispatch.call(workbooks, "Open", inputFile).toDispatch();
            Dispatch.call(excel, "ExportAsFixedFormat", EXCEL_FORMAT_PDF, pdfFile);
        } catch (Throwable t) {
            throw new IOException("excel转pdf失败", t);
        } finally {
            if (excel != null) {
                Dispatch.call(excel, "Close", false);
            }
            if (app != null) {
                app.invoke("Quit");
            }
            // 关闭WINWORD.exe进程
            ComThread.Release();
        }
    }

    public static void pptToPdfByMsOffice(String inputFile, String pdfFile) throws IOException {
        pptToPdf(inputFile, pdfFile, PROGRAM_PPT_OFFICE);
    }

    public static void pptToPdfByWps(String inputFile, String pdfFile) throws IOException {
        pptToPdf(inputFile, pdfFile, PROGRAM_PPT_WPS);
    }

    private static void pptToPdf(String inputFile, String pdfFile, String program) throws IOException {
        ActiveXComponent app = null;
        Dispatch ppt = null;
        try {
            // 创建一个ppt对象
            app = new ActiveXComponent(program);
            // 不能设置Visible=false。Hiding the application window is not allowed.
            // 禁用宏
            app.setProperty("AutomationSecurity", new Variant(3));
            // 获取文挡属性
            Dispatch presentations = app.getProperty("Presentations").toDispatch();
            // 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
            // 文档：https://learn.microsoft.com/zh-cn/office/vba/api/powerpoint.presentations.open
            ppt = Dispatch.call(presentations, "Open", inputFile,
                    true, // ReadOnly
                    false, // Untitled指定文件是否有标题
                    false // WithWindow指定文件是否可见
            ).toDispatch();
            Dispatch.call(ppt, "SaveAs", pdfFile, PPT_FORMAT_PDF);
        } catch (Throwable t) {
            throw new IOException("ppt转换pdf失败", t);
        } finally {
            if (ppt != null) {
                Dispatch.call(ppt, "Close");
            }
            if (app != null) {
                app.invoke("Quit");
            }
            // 关闭WINWORD.exe进程
            ComThread.Release();
        }
    }
}
```

## 二、Linux环境调用LibreOffice

### 安装LibreOffice
1. 去[LibreOffice下载页](https://zh-cn.libreoffice.org/download/libreoffice/)下载对应Linux发行版的LibreOffice。
2. 解压后安装。
    ```shell
    sudo dpkg -i DEBS/*.deb
    ```
    或者
    ```shell
    sudo yum install /RPMS/*.rpm
    ```
3. #### 增加中文字体，不然转换时中文会乱码。
    ```shell
    # 检查系统中是否有中文字体
    fc-list :lang=zh
    ```
    如果返回空，请按照[这个文档](https://zh-cn.libreoffice.org/download/fonts/)去安装中文字体。
4. 安装完成之后运行一下`libreoffice7.3 --help`，可能会提示库找不到。我在CentOS7系统安装时就遇到了libcairo.so.2，libcups.so.2，libSM.so.6这三个库找不到的问题，执行下面几条命令安装需要的库：
    ```shell
    yum install cairo -y
    yum install cups-libs -y
    yum install libSM -y
    ```
    如果遇到其他库找不到的问题可以参考上面解决。

### 转换代码：
```java
public class LibreOfficeToPdfUtil {
    public static void toPdf(String inputFile, String pdfFile) throws BootstrapException, Exception {
        XComponentContext context = Bootstrap.bootstrap();
        XMultiComponentFactory serviceManager = context.getServiceManager();
        Object desktop = serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context);
        XDesktop xDesktop = UnoRuntime.queryInterface(XDesktop.class, desktop);

        XComponentLoader xComponentLoader = UnoRuntime.queryInterface(XComponentLoader.class, xDesktop);
        PropertyValue hidden = new PropertyValue();
        hidden.Name = "Hidden";
        hidden.Value = Boolean.TRUE;
        XComponent xComponent = xComponentLoader.loadComponentFromURL("file:///" + inputFile, "_blank", 0, new PropertyValue[]{hidden});


        XStorable xStorable = UnoRuntime.queryInterface(XStorable.class, xComponent);
        PropertyValue overwrite = new PropertyValue();
        overwrite.Name = "Overwrite";
        overwrite.Value = Boolean.TRUE;
        PropertyValue filterName = new PropertyValue();
        filterName.Name = "FilterName";
        filterName.Value = "writer_pdf_Export";
        xStorable.storeToURL("file:///" + pdfFile, new PropertyValue[]{overwrite, filterName});


        xDesktop.terminate();
    }
}
```

### 运行时报错解决
1. com.sun.star.comp.helper.BootstrapException: no office executable found!
   **解决办法**：添加java参数`-Xbootclasspath/a:/opt/libreoffice7.3/program/`
   **原因**：启动时会从classpath中寻找soffice文件，所以得把soffice所在目录加到classpath中。
   > 可以通过<code>readlink &#96;which libreoffice7.3&#96;</code>看libreoffice安装在哪个目录。
   
2. 点运行后一直等待，5分钟后抛出BootstrapException异常
   **原因**：libreoffice进程启动失败。
   **解决办法**：可以在命令行运行一下libreoffice看看是否缺少库，然后参考上面[安装步骤](#安装libreoffice)解决。

3. 转换之后的pdf中文显示方框
   **原因**：系统中缺少中文字体。
   **解决办法**：参考[安装步骤](#安装libreoffice)添加中文字体。
   

参考文档：
1. [linux安装libreOffice](https://www.cnblogs.com/liangbo-/p/11424292.html)
2. [Convert Microsoft Word to PDF - using Java and LibreOffice (UNO API)](https://www.codeproject.com/Tips/988667/Convert-Microsoft-Word-to-PDF-using-Java-and-Libre)
3. [Java中常用的几种DOCX转PDF方法](https://segmentfault.com/a/1190000006789644)