package replacemark;

import org.apache.poi.xwpf.usermodel.*;
import org.springframework.util.StringUtils;

import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 *  替换文档中的段落和表格占位符
 * @author corey
 * @version 1.0
 * @date 2020/5/9 9:14 上午
 */
public class ReplaceUtil {

    /**
     * 替换段落中的占位符
     * @param doc 需要替换的文档
     * @param params 替换的参数，key=占位符，value=实际值
     */
    public static void replaceInPara(XWPFDocument doc, Map<String,Object> params)  {
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext()) {
            para = iterator.next();
            if(!StringUtils.isEmpty(para.getParagraphText())){
                replaceInPara(para, params);
            }
        }
    }

    /**
     * 替换段落中的占位符
     * @param para
     */
    public static void replaceInPara(XWPFParagraph para, Map<String,Object> params)  {
        // 获取当前段落的文本
        String sourceText = para.getParagraphText();
        // 控制变量
        boolean replace = false;
        for (Map.Entry<String, Object> entry : params.entrySet()) {
            String key = entry.getKey();
            if(sourceText.indexOf(key)!=-1){
                Object value = entry.getValue();
                if(value instanceof String){
                    // 替换文本占位符
                    sourceText = sourceText.replace(key, value.toString());
                    replace = true;
                }
            }
        }
        if(replace){
            // 获取段落中的行数
            List<XWPFRun> runList = para.getRuns();
            for (int i=runList.size();i>=0;i--){
                // 删除之前的行
                para.removeRun(i);
            }
            // 创建一个新的文本并设置为替换后的值 这样操作之后之前文本的样式就没有了，待改进
            para.createRun().setText(sourceText);
        }
    }

    /**
     * 替换表格中的占位符
     * @param doc
     * @param params
     */
    public static void replaceTable(XWPFDocument doc,Map<String,Object> params){
        // 获取文档中所有的表格
        Iterator<XWPFTable> iterator = doc.getTablesIterator();
        XWPFTable table;
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;
        while (iterator.hasNext()) {
            table = iterator.next();
            if (table.getRows().size() > 1) {
                //判断表格是需要替换还是需要插入，判断逻辑有${为替换，
                if (matcher(table.getText()).find()) {
                    rows = table.getRows();
                    for (XWPFTableRow row : rows) {
                        cells = row.getTableCells();
                        for (XWPFTableCell cell : cells) {
                            paras = cell.getParagraphs();
                            for (XWPFParagraph para : paras) {
                                replaceInPara(para, params);
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * 正则匹配字符串
     *
     * @param str
     * @return
     */
    private static Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }
}
