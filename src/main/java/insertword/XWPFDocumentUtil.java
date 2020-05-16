package insertword;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.util.Map;

/**
 * @author corey
 * @version 1.0
 * @date 2020/5/5 2:04 下午
 */
public class XWPFDocumentUtil {


    /**
     * 往doc的标记位置插入富文本内容 注意：目前支持富文本里面带url的图片，不支持base64编码的图片
     * @param doc 需要插入内容的Word
     * @param ritchtextMap 标记位置对应的富文本内容
     * @param watermark 水印内容
     */
    public static void wordInsertRitchText(XWPFDocument doc,Map<String, String> ritchtextMap,String watermark) {
        try {
            long beginTime = System.currentTimeMillis();
            // 如果需要替换多份富文本，通过Map来操作，key:要替换的标记，value：要替换的富文本内容
            for(Map.Entry<String,String> entry: ritchtextMap.entrySet()){
                for (XWPFParagraph paragraph : doc.getParagraphs()) {
                    if(entry.getKey().equals(paragraph.getText().trim())){
                        // 在标记处插入指定富文本内容
                        HtmlUtil.resolveHtml(entry.getValue(),doc,paragraph);
                        // 删除要被替换的标记
                        doc.removeBodyElement(doc.getPosOfParagraph(paragraph));
                        break;
                    }
                }
            }
            // 添加水印
            WatermarkUtil.waterMarkDocXDocument(doc,watermark==null?"create by corey":watermark);
            // 设置目录 待开发
            System.out.println("生成成功!,一共耗时"+(System.currentTimeMillis()-beginTime)+"毫秒");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
