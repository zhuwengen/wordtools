package insertword.test;

import insertword.HtmlUtil;
import insertword.XWPFDocumentUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

/**
 * @author corey
 * @version 1.0
 * @date 2020/5/5 2:04 下午
 */
public class XWPFDocumentTest {


    public static void main(String[] args) {
        InputStream in = null;
        // 插入富文本后Word的输出目录 这个文件会自动生成
        OutputStream out = null;
        try {
            out = new FileOutputStream("/Users/corey/Desktop/temp/wordtools/合并文档1.docx");;
            // 模拟需要被插入的word 请求doc目录下的insetritchsource.docx放到自己合适的目录，并修改此处的路径
            String mainFilePath = "/Users/corey/Desktop/temp/wordtools/insetritchsource.docx";
            // 需要插入的word文件
            File mainfile = new File(mainFilePath);
            in = new FileInputStream(mainfile);
            OPCPackage srcPackage = OPCPackage.open(in);
            XWPFDocument doc = new XWPFDocument(srcPackage);
            // 实际插入富文本的逻辑 同时打上水印
            XWPFDocumentUtil.wordInsertRitchText(doc,insertRitch(),"这里写水印内容");
            // 将doc输出到指定目录
            doc.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            HtmlUtil.closeStream(in,out);
        }
    }

    /**
     * 模拟富文本占位符需要替换的内容
     * @return
     */
    public static Map<String,String> insertRitch(){
        // 模拟富文本内容 这里需要特别注意富文本中的图片地址，可以使用doc目录下的图片和富文本内容，切记替换富文本中的图片地址，
        // 请求doc目录下的富文本输出内容.txt放到自己合适的目录，并修改此处的路径
        String sourceFileName = "/Users/corey/Desktop/temp/wordtools/富文本输出内容.txt";
        File file = new File(sourceFileName);
        String content = HtmlUtil.txt2String(file);
        // 模拟标记位置及对应的富文本内容
        Map<String, String> ritchtextMap = new HashMap<String,String>();
        ritchtextMap.put("${ritchtext}", content);
        ritchtextMap.put("${ritchtext1}", content);
        return ritchtextMap;
    }

}
