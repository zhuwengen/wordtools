package replacemark.test;

import insertword.HtmlUtil;
import insertword.XWPFDocumentUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import replacemark.ReplaceUtil;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

/**
 * 注意事项请参考: XWPFDocumentTest.java
 * 功能描述：可以替换文档中段落占位符、替换文档表格中占位符、插入富文本内容、添加水印
 * @author corey
 * @version 1.0
 * @date 2020/5/9 9:38 上午
 */
public class MakeWordTest {
    public static void main(String[] args) throws Exception {
        // 需要替换的文档路径
        String mainFilePath = "/Users/corey/Desktop/temp/wordtools/replacesource.docx";
        File mainFile = new File(mainFilePath);
        InputStream in = new FileInputStream(mainFile);
        OPCPackage srcPackage = OPCPackage.open(in);
        XWPFDocument doc = new XWPFDocument(srcPackage);
        // 替换文档中段落的占位符
        ReplaceUtil.replaceInPara(doc,createParaParamsMap());
        // 替换文档中表格里面的占位符
        ReplaceUtil.replaceTable(doc,createTableParamsMap());
        // 插入富文本框到文本中指定的占位符
        XWPFDocumentUtil.wordInsertRitchText(doc,insertRitch(),"这里是水印");
        // 插入富文本后Word的输出目录
        OutputStream dest = new FileOutputStream("/Users/corey/Desktop/temp/wordtools/合并文档2.docx");
        doc.write(dest);
        // 关闭流
        HtmlUtil.closeStream(in,dest);
    }



    /**
     * 创建文本占位符需要替换的内容
     * @return
     */
    public static Map<String,Object> createParaParamsMap(){
        Map<String, Object> map = new HashMap<>();
        map.put("${page1_projectName}", "长城贴瓷砖项目");
        map.put("${page1_custName}", "宏伟装饰有限公司");
        map.put("${page1_leaseAmount}", "100，000，000");
        map.put("${page1_guarName}", "直售");
        map.put("${page1_update}", "2020-5-10");
        map.put("${page1_applyTime}", "2019-12-12");
        map.put("${page1_mainManagerName}", "张三");
        map.put("${page1_assistManagerName}", "李四");
        map.put("${page1_surveyTime}", "2019-12-22");
        map.put("${page1_fieldTime}", "2020-1-23");
        map.put("${page1_count}", "1");
        map.put("${page1_headName}", "王五");
        return map;
    }

    /**
     * 创建表格占位符需要替换的内容
     * @return
     */
    public static Map<String,Object> createTableParamsMap(){
        Map<String, Object> map = new HashMap<>();
        map.put("${table1_custName}", "宏伟装饰有限公司");
        map.put("${table1_projectName}", "长城贴瓷砖项目");
        map.put("${table1_xingzhi}", "直售");
        map.put("${table1_fangshi}", "直售");
        return map;
    }

    /**
     * 模拟富文本占位符需要替换的内容
     * @return
     */
    public static Map<String,String> insertRitch(){
        // 模拟富文本内容
        String sourceFileName = "/Users/corey/Desktop/temp/wordtools/富文本输出内容.txt";
        File file = new File(sourceFileName);
        String content = HtmlUtil.txt2String(file);
        // 模拟标记位置及对应的富文本内容
        Map<String, String> ritchtextMap = new HashMap<>();
        ritchtextMap.put("${ritchtext}", content);
        ritchtextMap.put("${ritchtext1}", content);
        return ritchtextMap;
    }
}
