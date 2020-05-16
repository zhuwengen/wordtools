package insertword;

/**
 * html 元素枚举映射类
 * @author corey
 * @version 1.0
 * @date 2020/5/5 9:36 下午
 */
public enum ElementEnum {
    H1("h1","h1","一级标题"),
    H2("h2","h2","二级标题"),
    H3("h3","h3","三级标题"),
    P("p", "paragraph", "段落"),
    STRONG("strong","","加粗"),
    I("i","","斜体"),
    U("u", "", "字体下划线"),
    // IMG("img", "imgbase64", "base64图片"),
    IMG_URL("img", "imgurl", "图片url"),
    TABLE("table","table","表格")

    ;

    private String code;
    private String value;
    private String desc;

    public String getCode() {
        return code;
    }

    public String getValue() {
        return value;
    }

    public String getDesc() {
        return desc;
    }

    ElementEnum(String code, String value, String desc) {
        this.code = code;
        this.value = value;
        this.desc = desc;
    }

    public static String getValueByCode(String code) {
        for (ElementEnum e : ElementEnum.values()) {
            if (e.getCode().equalsIgnoreCase(code)) {
                return e.getValue();
            }
        }
        return null;
    }
}
