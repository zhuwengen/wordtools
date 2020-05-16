package magerword;


import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.springframework.util.ObjectUtils;
import org.springframework.util.StringUtils;

/**
 * Html工具类
 */
public class HtmlUtil {

    /**
     * 给document添加指定元素
     * @param document
     */
    public static void addElement(Document document){
        if(ObjectUtils.isEmpty(document)){
            throw new NullPointerException("不允许为空的对象添加元素");
        }
        Elements elements = document.getAllElements();
        for(Element e:elements){
            String attrName = ElementEnum.getValueByCode(e.tag().getName());
            if(!StringUtils.isEmpty(attrName)) {
                e.attr(CommonConStant.COMMONATTR, attrName);
            }
        }
    }
}
