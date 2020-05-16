package magerword;

import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;
import org.docx4j.wml.PPrBase.Ind;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import sun.misc.BASE64Decoder;

import java.io.*;
import java.math.BigInteger;
import java.util.List;

/**
 * @program: htmltoword
 * @description: html 转 docx
 * @author: corey
 * @create: 2020-04-29 14:10
 **/
public class HtmlToWord {
    private static ObjectFactory factory;
    private static WordprocessingMLPackage wordMLPackage;

    /**
     * 将一段富文本字符串转为一个字节数组
     * @param data
     * @return
     */
    public static byte[] resolveHtml(String data) {
        Document document = Jsoup.parseBodyFragment(data, "UTF-8");
        ByteArrayOutputStream out = null;
        try {
            wordMLPackage = WordprocessingMLPackage.createPackage();
            factory = Context.getWmlObjectFactory();
            Relationship relationship = createFooterPart();
            createFooterReference(relationship);
            MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
            alterStyleSheet();
            // 添加固定元素
            HtmlUtil.addElement(document);
            Elements elements = document.select("["+CommonConStant.COMMONATTR+"]");
            for (Element em : elements) {
                switch (em.attr(CommonConStant.COMMONATTR)) {
                    case "title":
                        documentPart.addStyledParagraphOfText("Title", em.text());
                        break;
                    case "subtitle":
                        documentPart.addStyledParagraphOfText("Subtitle", em.text());
                        break;
                    case "imgurl":
                        String imgSrc = em.attr("src");
                        File file = new File(imgSrc);
                        byte[] bytes = convertImageToByteArray(file);
                        addImageToPackage(wordMLPackage, bytes);
                        break;
                    case "imgbase64":
                        String base64chart = em.attr("src");
                        String base64 = base64chart.substring(base64chart.indexOf(";base64,") + 8);
                        addImageToPackage(wordMLPackage, new BASE64Decoder().decodeBuffer(base64));
                        break;
                    case "table":
                        Tbl table = addTable(em);
                        documentPart.addObject(table);
                        break;
                    case "h1":
                        documentPart.addStyledParagraphOfText("Heading1", em.text());
                        break;
                    case "h2":
                        documentPart.addStyledParagraphOfText("Heading2", em.text());
                        break;
                    case "h3":
                        documentPart.addStyledParagraphOfText("Heading3", em.text());
                        break;
                    case "paragraph":
                        P p = addParapraph(em.text());
                        //设置首行缩进
                        setFirstLine(p,"400");
                        documentPart.getContent().add(p);
                        break;
                    default:
                        documentPart.addParagraphOfText(em.text());
                        break;
                }
            }
            addPageBreak(documentPart);
             out = new ByteArrayOutputStream();
            wordMLPackage.save(out);
            return out.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }finally {
            if(out!=null){
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * @Title: addParapraph
     * @Description: (文本转段落)
     * @param @param text
     * @param @return    设定文件
     * @return P    返回类型
     * @throws
     */
    private static P addParapraph(String text) {
        factory=Context.getWmlObjectFactory();
        P paragraph = factory.createP();
        Text t = factory.createText();
        t.setValue(text);
        R run = factory.createR();
        run.getContent().add(t);
        paragraph.getContent().add(run);
        RPr runProperties = factory.createRPr();
        run.setRPr(runProperties);
        return paragraph;
    }

    /**
     * @Title: setFirstLine
     * @param @param p
     * @param @param str    设定文件
     * @return void    返回类型
     * @throws
     */
    private static void setFirstLine(P p ,String str) {
        PPr ppr = getPPr(p);
        Ind ind = ppr.getInd();
        if (ind == null) {
            ind = new Ind();
            ppr.setInd(ind);
        }
        ind.setFirstLine(new BigInteger(str));
    };

    private static PPr getPPr(P p) {
        PPr ppr = p.getPPr();
        if (ppr == null) {
            ppr = new PPr();
            p.setPPr(ppr);
        }
        return ppr;
    }

    /**
     *         table @param @return 设定文件 @return Tbl 返回类型 @throws
     */
    private static Tbl addTable(Element table) {
        factory = Context.getWmlObjectFactory();
        Tbl tbl = factory.createTbl();
        addBorders(tbl);
        Elements trs = table.getElementsByTag("tr");
        for (Element tr : trs) {
            Tr fTr = addTableTr(tr);
            tbl.getContent().add(fTr);
        }
        return tbl;
    }

    /**
     *         tr @param @return 设定文件 @return Tr 返回类型 @throws
     */
    private static Tr addTableTr(Element tr) {
        Elements tds = tr.getElementsByTag("th").isEmpty() ? tr.getElementsByTag("td") : tr.getElementsByTag("th");
        Tr ftr = factory.createTr();
        for (int i = 0, j = tds.size(); i < j; i++) {
            Tc ftd = factory.createTc();
            setCellWidth(ftd, 1000);
            ftd.getContent().add(wordMLPackage.getMainDocumentPart().createParagraphOfText(tds.get(i).text()));
            ftr.getContent().add(ftd);
        }
        return ftr;
    }

    /**
     * 本方法创建一个单元格属性集对象和一个表格宽度对象. 将给定的宽度设置到宽度对象然后将其添加到 属性集对象. 最后将属性集对象设置到单元格中.
     */
    private static void setCellWidth(Tc tableCell, int width) {
        TcPr tableCellProperties = new TcPr();
        TblWidth tableWidth = new TblWidth();
        tableWidth.setW(BigInteger.valueOf(width));
        tableCellProperties.setTcW(tableWidth);
        tableCell.setTcPr(tableCellProperties);
    }

    /**
     * 本方法为表格添加边框
     */
    private static void addBorders(Tbl table) {
        table.setTblPr(new TblPr());
        CTBorder border = new CTBorder();
        border.setColor("auto");
        border.setSz(new BigInteger("4"));
        border.setSpace(new BigInteger("0"));
        border.setVal(STBorder.SINGLE);

        TblBorders borders = new TblBorders();
        borders.setBottom(border);
        borders.setLeft(border);
        borders.setRight(border);
        borders.setTop(border);
        borders.setInsideH(border);
        borders.setInsideV(border);
        table.getTblPr().setTblBorders(borders);
    }

    /**
     * 将图片从文件对象转换成字节数组.
     *
     * @param file
     *            将要转换的文件
     * @return 包含图片字节数据的字节数组
     * @throws FileNotFoundException
     * @throws IOException
     */
    private static byte[] convertImageToByteArray(File file) throws FileNotFoundException, IOException {
        InputStream is = new FileInputStream(file);
        long length = file.length();
        // 不能使用long类型创建数组, 需要用int类型.
        if (length > Integer.MAX_VALUE) {
            System.out.println("File too large!!");
        }
        byte[] bytes = new byte[(int) length];
        int offset = 0;
        int numRead = 0;
        while (offset < bytes.length && (numRead = is.read(bytes, offset, bytes.length - offset)) >= 0) {
            offset += numRead;
        }
        // 确认所有的字节都没读取
        if (offset < bytes.length) {
            System.out.println("Could not completely read file " + file.getName());
        }
        is.close();
        return bytes;
    }

    /**
     * Docx4j拥有一个由字节数组创建图片部件的工具方法, 随后将其添加到给定的包中. 为了能将图片添加 到一个段落中, 我们需要将图片转换成内联对象.
     * 这也有一个方法, 方法需要文件名提示, 替换文本, 两个id标识符和一个是嵌入还是链接到的指示作为参数. 一个id用于文档中绘图对象不可见的属性,
     * 另一个id用于图片本身不可见的绘制属性. 最后我们将内联 对象添加到段落中并将段落添加到包的主文档部件.
     *
     * @param wordMLPackage
     *            要添加图片的包
     * @param bytes
     *            图片对应的字节数组
     * @throws Exception
     *             不幸的createImageInline方法抛出一个异常(没有更多具体的异常类型)
     */
    private static void addImageToPackage(WordprocessingMLPackage wordMLPackage, byte[] bytes) throws Exception {
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);

        int docPrId = 1;
        int cNvPrId = 2;
        Inline inline = imagePart.createImageInline("Filename hint", "Alternative text", docPrId, cNvPrId, false);

        P paragraph = addInlineImageToParagraph(inline);

        wordMLPackage.getMainDocumentPart().addObject(paragraph);
    }

    /**
     * 创建一个对象工厂并用它创建一个段落和一个可运行块R. 然后将可运行块添加到段落中. 接下来创建一个图画并将其添加到可运行块R中. 最后我们将内联
     * 对象添加到图画中并返回段落对象.
     *
     * @param inline
     *            包含图片的内联对象.
     * @return 包含图片的段落
     */
    private static P addInlineImageToParagraph(Inline inline) {
        // 添加内联对象到一个段落中
        ObjectFactory factory = new ObjectFactory();
        P paragraph = factory.createP();
        R run = factory.createR();
        paragraph.getContent().add(run);
        Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        return paragraph;
    }

    /**
     * This method alters the default style sheet that is part of each document.
     *
     * To do this, we first retrieve the style sheet from the package and then get
     * the Styles object from it. From this object, we get the list of actual styles
     * and iterate over them. We check against all styles we want to alter and apply
     * the alterations if applicable.
     *
     * @param
     */
    public static void alterStyleSheet() {
        StyleDefinitionsPart styleDefinitionsPart = wordMLPackage.getMainDocumentPart().getStyleDefinitionsPart();
        Styles styles = null;
        try {
            styles = styleDefinitionsPart.getContents();
        } catch (Docx4JException e) {
            e.printStackTrace();
        }

        List<Style> stylesList = styles.getStyle();
        for (Style style : stylesList) {
            if (style.getStyleId().equals("Normal")) {
                alterNormalStyle(style);
            } else if (style.getStyleId().equals("Heading1")) {
                alterHeading1Style(style);
            } else if (style.getStyleId().equals("Heading2")) {
                alterHeading2Style(style);
            } else if (style.getStyleId().equals("Title") || style.getStyleId().equals("Subtitle")) {
                getRunPropertiesAndRemoveThemeInfo(style);
            }
        }
    }

    /**
     * First we create a run properties object as we want to remove nearly all of
     * the existing styling. Then we change the font and font size and set the run
     * properties on the given style. As in previous examples, the font size is
     * defined to be in half-point size.
     */
    private static void alterNormalStyle(Style style) {
        // we want to change (or remove) almost all the run properties of the
        // normal style, so we create a new one.
        RPr rpr = new RPr();
        changeFontToArial(rpr);
        changeFontSize(rpr, 20);
        style.setRPr(rpr);
    }

    /**
     * For this style, we get the existing run properties from the style and remove
     * the theme font information from them. Then we also remove the bold styling,
     * change the font size (half-points) and add an underline.
     */
    private static void alterHeading1Style(Style style) {
        RPr rpr = getRunPropertiesAndRemoveThemeInfo(style);
        removeBoldStyle(rpr);
        changeFontSize(rpr, 28);
        /* addUnderline(rpr); */
    }

    private static void alterHeading2Style(Style style) {
        RPr rpr = getRunPropertiesAndRemoveThemeInfo(style);
        removeBoldStyle(rpr);
        changeFontSize(rpr, 24);

        /* addUnderline(rpr); */
    }

    private static RPr getRunPropertiesAndRemoveThemeInfo(Style style) {
        // We only want to change some settings, so we get the existing run
        // properties from the style.
        RPr rpr = style.getRPr();
        removeThemeFontInformation(rpr);
        return rpr;
    }

    /**
     * Change the font of the given run properties to Arial.
     *
     * A run font specifies the fonts which shall be used to display the contents of
     * the run. Of the four possible types of content, we change the styling of two
     * of them: ASCII and High ANSI. Finally we add the run font to the run
     * properties.
     *
     * @param runProperties
     */
    private static void changeFontToArial(RPr runProperties) {
        RFonts runFont = new RFonts();
        runFont.setAscii("Arial");
        runFont.setHAnsi("Arial");
        runProperties.setRFonts(runFont);
    }

    /**
     * Change the font size of the given run properties to the given value.
     *
     * @param runProperties
     * @param fontSize
     *            Twice the size needed, as it is specified as half-point value
     */
    private static void changeFontSize(RPr runProperties, int fontSize) {
        HpsMeasure size = new HpsMeasure();
        size.setVal(BigInteger.valueOf(fontSize));
        runProperties.setSz(size);
    }

    /**
     * Removes the theme font information from the run properties. If this is not
     * removed then the styles based on the normal style won't inherit the Arial
     * font from the normal style.
     *
     * @param runProperties
     */
    private static void removeThemeFontInformation(RPr runProperties) {
        runProperties.getRFonts().setAsciiTheme(null);
        runProperties.getRFonts().setHAnsiTheme(null);
    }

    /**
     * Removes the Bold styling from the run properties.
     *
     * @param runProperties
     */
    private static void removeBoldStyle(RPr runProperties) {
        runProperties.getB().setVal(false);
    }



    /**
     * As in the previous example, this method creates a footer part and adds it to
     * the main document and then returns the corresponding relationship.
     *
     * @return
     * @throws InvalidFormatException
     */
    private static Relationship createFooterPart() throws InvalidFormatException {
        FooterPart footerPart = new FooterPart();
        footerPart.setPackage(wordMLPackage);

        footerPart.setJaxbElement(createFooterWithPageNr());

        return wordMLPackage.getMainDocumentPart().addTargetPart(footerPart);
    }

    /**
     * As in the previous example, we create a footer and a paragraph object. But
     * this time, instead of adding text to a run, we add a field. And just as with
     * the table of content, we have to add a begin and end character around the
     * actual field with the page number. Finally we add the paragraph to the
     * content of the footer and then return it.
     *
     * @return
     */
    public static Ftr createFooterWithPageNr() {
        Ftr ftr = factory.createFtr();
        P paragraph = factory.createP();

        addFieldBegin(paragraph);
        addPageNumberField(paragraph);
        addFieldEnd(paragraph);

        ftr.getContent().add(paragraph);
        return ftr;
    }

    /**
     * Creating the page number field is nearly the same as creating the field in
     * the TOC example. The only difference is in the value. We use the PAGE
     * command, which prints the number of the current page, together with the
     * MERGEFORMAT switch, which indicates that the current formatting should be
     * preserved when the field is updated.
     *
     * @param paragraph
     */
    private static void addPageNumberField(P paragraph) {
        R run = factory.createR();
        Text txt = new Text();
        txt.setSpace("preserve");
        txt.setValue(" PAGE   \\* MERGEFORMAT ");
        run.getContent().add(factory.createRInstrText(txt));
        paragraph.getContent().add(run);
    }

    /**
     * Every fields needs to be delimited by complex field characters. This method
     * adds the delimiter that precedes the actual field to the given paragraph.
     *
     * @param paragraph
     */
    private static void addFieldBegin(P paragraph) {
        R run = factory.createR();
        FldChar fldchar = factory.createFldChar();
        fldchar.setFldCharType(STFldCharType.BEGIN);
        run.getContent().add(fldchar);
        paragraph.getContent().add(run);
    }

    /**
     * Every fields needs to be delimited by complex field characters. This method
     * adds the delimiter that follows the actual field to the given paragraph.
     *
     * @param paragraph
     */
    private static void addFieldEnd(P paragraph) {
        FldChar fldcharend = factory.createFldChar();
        fldcharend.setFldCharType(STFldCharType.END);
        R run3 = factory.createR();
        run3.getContent().add(fldcharend);
        paragraph.getContent().add(run3);
    }

    /**
     * This method fetches the document final section properties, and adds a newly
     * created footer reference to them.
     *
     * @param relationship
     */
    public static void createFooterReference(Relationship relationship) {

        List<SectionWrapper> sections = wordMLPackage.getDocumentModel().getSections();

        SectPr sectPr = sections.get(sections.size() - 1).getSectPr();
        // There is always a section wrapper, but it might not contain a sectPr
        if (sectPr == null) {
            sectPr = factory.createSectPr();
            wordMLPackage.getMainDocumentPart().addObject(sectPr);
            sections.get(sections.size() - 1).setSectPr(sectPr);
        }

        FooterReference footerReference = factory.createFooterReference();
        footerReference.setId(relationship.getId());
        footerReference.setType(HdrFtrRef.DEFAULT);
        sectPr.getEGHdrFtrReferences().add(footerReference);
    }

    /**
     * Adds a page break to the document.
     *
     * @param documentPart
     */
    private static void addPageBreak(MainDocumentPart documentPart) {
        Br breakObj = new Br();
        breakObj.setType(STBrType.PAGE);

        P paragraph = factory.createP();
        paragraph.getContent().add(breakObj);
        try {
            documentPart.getContents().getBody().getContent().add(paragraph);
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
    }
}