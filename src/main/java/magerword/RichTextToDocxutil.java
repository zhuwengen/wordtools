package magerword;

import java.io.*;

/**
 * @author corey
 * @version 1.0
 * @date 2020/5/3 5:51 下午
 * @Desc 将富文本txt导出成word
 */
public class RichTextToDocxutil {
    /**
     * 导出富本框到docx
     */
    public static void outRichTextToDocx(String sourceFilePath ,String outFilePath) {
        File file = new File(sourceFilePath);
        String content = txt2String(file);
        OutputStream out = null;
        try {
            // 输入富文本内容，返回字节数组
            byte[] result = HtmlToWord.resolveHtml(content);
            //输出文件
            out = new FileOutputStream(outFilePath);
            out.write(result);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }

        }
    }

    /**
     * 读取txt文件的内容
     *
     * @param file 想要读取的文件对象
     * @return 返回文件内容
     */
    public static String txt2String(File file) {
        StringBuilder result = new StringBuilder();
        try {
            // 构造一个BufferedReader类来读取文件
            BufferedReader br = new BufferedReader(new FileReader(file));
            String s = null;
            // 使用readLine方法，一次读一行
            while ((s = br.readLine()) != null) {
                result.append(System.lineSeparator() + s);
            }
            br.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result.toString();
    }
}
