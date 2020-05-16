package magerword;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.springframework.util.ObjectUtils;

import java.io.*;
import java.util.*;

/**
 * @program: 合并多份word文件
 * @description:
 * @author: corey
 * @create: 2020-04-29 19:04
 **/
public class MagerUtil {
    /**
     * 合并多个Word
     * @param filepaths
     * @throws Exception
     */
    public static void mergeDoc(String... filepaths) throws Exception {
        // 需要配置导出文件路径 记得替换为自己电脑的路径
        OutputStream dest = new FileOutputStream("/Users/corey/Desktop/temp/wordtools/合并文档3.docx");
        List<CTBody> ctBodyList = new ArrayList<>();
        List<XWPFDocument> srcDocuments = new ArrayList<>();
        for (String filepath : filepaths) {
            InputStream in = null;
            OPCPackage srcPackage = null;
            try {
                in = new FileInputStream(filepath);
                srcPackage = OPCPackage.open(in);
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                closeStream(in);
            }
            XWPFDocument srcDocument = new XWPFDocument(srcPackage);
            CTBody srcBody = srcDocument.getDocument().getBody();
            ctBodyList.add(srcBody);
            srcDocuments.add(srcDocument);
        }
        if (!ObjectUtils.isEmpty(ctBodyList)) {
            appendBody(ctBodyList);
            srcDocuments.get(0).write(dest);
        }
    }

    /**
     * 拼接所有的文档元素
     * @param ctBodyList
     * @throws Exception
     */
    private static void appendBody(List<CTBody> ctBodyList) throws Exception {
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        // 所有的xmlns
        StringBuffer allAmlns = new StringBuffer();
        // 所有文档的内部元素
        StringBuffer allElement = new StringBuffer();
        ctBodyList.forEach(ct -> {
            // 拿到每一个文档的完整xml
            String appentString = ct.xmlText();
            // 拼接所有的xmlns
            allAmlns.append(appentString.substring(appentString.indexOf("xmlns"), appentString.indexOf(">")));
            // 拼接所有的内部元素
            allElement.append(appentString.substring(appentString.indexOf(">") + 1, appentString.lastIndexOf("</")));
        });
        // 将xmlns去重
        String distinctPrefix = distinctXmlns(allAmlns.toString());
        // 合并文档
        CTBody makeBody = CTBody.Factory.parse(distinctPrefix + allElement.toString() + "</xml-fragment>");
        ctBodyList.get(0).set(makeBody);
    }

    /**
     * 去重合并xml的Xmlns
     *
     * @param prefix
     * @return
     */
    public static String distinctXmlns(String prefix) {
        int start = prefix.indexOf("xmlns");
        int end = prefix.indexOf("xmlns", start + 1);
        Set s = new HashSet();
        while (end > 0) {
            s.add(prefix.substring(start, end));
            start = end;
            end = prefix.indexOf("xmlns", start + 1);
        }
        String xmlHead = "<xml-fragment ";
        StringBuffer sb = new StringBuffer(xmlHead);
        Map<String, String> map = distinctXmlns(s);
        for (Map.Entry<String, String> entry : map.entrySet()) {
            sb.append(" ");
            sb.append(entry.getKey());
            sb.append("=");
            sb.append(entry.getValue());
        }
        sb.append(">");
        return sb.toString();
    }

    /**
     * xmlns 可能存在xmlns头相同但是指向地址不同的情况
     *
     * @param set
     * @return
     */
    public static Map<String, String> distinctXmlns(Set set) {
        Map<String, String> map = new HashMap();
        Iterator i = set.iterator();
        while (i.hasNext()) {
            String xmlns = (String) i.next();
            map.put(xmlns.substring(0, xmlns.indexOf("=")), xmlns.substring(xmlns.indexOf("=") + 1));
        }
        return map;
    }

    /**
     * 关闭流
     * 这一步可以放到公用工具类中，close的类型可以使用Closeable，这样就可以关闭input和output的流
     * @param inputStream
     */
    public static void closeStream(InputStream... inputStream) {
        for (InputStream i : inputStream) {
            if (i != null) {
                try {
                    i.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }
}