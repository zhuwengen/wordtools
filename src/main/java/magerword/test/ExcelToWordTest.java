package magerword.test;


import magerword.MagerUtil;
import magerword.RichTextToDocxutil;

/**
 * @author corey
 * @version 1.0
 * @date 2020/5/2 5:53 下午
 */
public class ExcelToWordTest {

    public static void main(String[] args) throws Exception{
        try {
            // 下面的路径都需要替换为自己的目录
            String sourceFilePath = "/Users/corey/Desktop/temp/wordtools/富文本输出内容.txt";
            String outFilePath = "/Users/corey/Desktop/temp/wordtools/导出富文本框.docx";
            String fileName1 = "/Users/corey/Desktop/temp/wordtools/导出富文本框.docx";
            String fileName2 = "/Users/corey/Desktop/temp/wordtools/导出富文本框.docx";
            String fileName3 = "/Users/corey/Desktop/temp/wordtools/导出富文本框.docx";
            // 将富文本字段转成docx
            RichTextToDocxutil.outRichTextToDocx(sourceFilePath,outFilePath);
            // 多文本合并
            MagerUtil.mergeDoc(fileName1,fileName2,fileName3);
            System.out.println("文档生成成功");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


}
