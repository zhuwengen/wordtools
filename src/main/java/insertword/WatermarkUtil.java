package insertword;

import com.microsoft.schemas.office.office.CTLock;
import com.microsoft.schemas.vml.*;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.util.stream.Stream;

/**
 * @desc 添加水印
 * @author corey
 * @version 1.0
 * @date 2020/5/5 10:07 下午
 */
public class WatermarkUtil {
    // word字体
    private static final String fontName = "宋体";
    // 字体大小
    private static final String fontSize = "0.2pt";
    // 字体颜色
    private static final String fontColor = "#d0d0d0";
    // 一个字平均长度，单位pt，用于：计算文本占用的长度（文本总个数*单字长度）
    private static  final Integer widthPerWord = 10;
    // 与顶部的间距
    private static Integer styleTop = 0;
    // 文本旋转角度
    private static final String styleRotation = "45";



    /**
     * 给文档添加水印
     * 此方法可以单独使用
     * @param doc
     * @param customText
     */
    public static void waterMarkDocXDocument(XWPFDocument doc,String customText){
        // 把整页都打上水印
        for (int lineIndex = -5; lineIndex < 20; lineIndex++) {
            styleTop = 100*lineIndex;
            waterMarkDocXDocument_0(doc,customText);
        }
    }



    /**
     * 为文档添加水印
     * @param doc 需要被处理的docx文档对象
     * @param customText 需要添加的水印文字
     */
    public static void waterMarkDocXDocument_0(XWPFDocument doc,String customText) {
        // 水印文字之间使用8个空格分隔
        customText = customText + repeatString(" ", 8);
        // 一行水印重复水印文字次数
        customText = repeatString(customText, 10);
        // 如果之前已经创建过 DEFAULT 的Header，将会复用
        XWPFHeader header = doc.createHeader(HeaderFooterType.DEFAULT);
        int size = header.getParagraphs().size();
        if (size == 0) {
            header.createParagraph();
        }
        CTP ctp = header.getParagraphArray(0).getCTP();
        byte[] rsidr = doc.getDocument().getBody().getPArray(0).getRsidR();
        byte[] rsidrdefault = doc.getDocument().getBody().getPArray(0).getRsidRDefault();
        ctp.setRsidP(rsidr);
        ctp.setRsidRDefault(rsidrdefault);
        CTPPr ppr = ctp.addNewPPr();
        ppr.addNewPStyle().setVal("Header");
        // 开始加水印
        CTR ctr = ctp.addNewR();
        CTRPr ctrpr = ctr.addNewRPr();
        ctrpr.addNewNoProof();
        CTGroup group = CTGroup.Factory.newInstance();
        CTShapetype shapetype = group.addNewShapetype();
        CTTextPath shapeTypeTextPath = shapetype.addNewTextpath();
        shapeTypeTextPath.setOn(STTrueFalse.T);
        shapeTypeTextPath.setFitshape(STTrueFalse.T);
        CTLock lock = shapetype.addNewLock();
        lock.setExt(STExt.VIEW);
        CTShape shape = group.addNewShape();
        shape.setId("PowerPlusWaterMarkObject");
        shape.setSpid("_x0000_s102");
        shape.setType("#_x0000_t136");
        // 设置形状样式（旋转，位置，相对路径等参数）
        shape.setStyle(getShapeStyle(customText));
        shape.setFillcolor(fontColor);
        // 字体设置为实心
        shape.setStroked(STTrueFalse.FALSE);
        // 绘制文本的路径
        CTTextPath shapeTextPath = shape.addNewTextpath();
        // 设置文本字体与大小
        shapeTextPath.setStyle("font-family:" + fontName + ";font-size:" + fontSize);
        shapeTextPath.setString(customText);
        CTPicture pict = ctr.addNewPict();
        pict.set(group);
    }

    /**
     * 构建Shape的样式参数
     * @param customText
     * @return
     */
    private static String getShapeStyle(String customText) {
        StringBuilder sb = new StringBuilder();
        // 文本path绘制的定位方式
        sb.append("position: ").append("absolute");
        // 计算文本占用的长度（文本总个数*单字长度）
        sb.append(";width: ").append(customText.length() * widthPerWord).append("pt");
        // 字体高度
        sb.append(";height: ").append("20pt");
        sb.append(";z-index: ").append("-251654144");
        sb.append(";mso-wrap-edited: ").append("f");
        // 设置水印的间隔，这是一个大坑，不能用top,必须要margin-top。
        sb.append(";margin-top: ").append(styleTop);
        sb.append(";mso-position-horizontal-relative: ").append("page");
        sb.append(";mso-position-vertical-relative: ").append("page");
        sb.append(";mso-position-vertical: ").append("left");
        sb.append(";mso-position-horizontal: ").append("center");
        sb.append(";rotation: ").append(styleRotation);
        return sb.toString();
    }

    /**
     * 将指定的字符串重复repeats次.
     */
    private static String repeatString(String pattern, int repeats) {
        StringBuilder buffer = new StringBuilder(pattern.length() * repeats);
        Stream.generate(() -> pattern).limit(repeats).forEach(buffer::append);
        return new String(buffer);
    }
}
