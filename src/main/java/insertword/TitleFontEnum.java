package insertword;

/**
 * @desc 设置标题字体大小
 * @author corey
 * @version 1.0
 * @date 2020/5/5 9:48 下午
 */
public enum TitleFontEnum {
    H1("h1", 26),
    H2("h2", 24),
    H3("h3", 22)
    ;
    private String title;
    private Integer font;

    public String getTitle() {
        return title;
    }

    public Integer getFont() {
        return font;
    }

    TitleFontEnum(String title, Integer font) {
        this.title = title;
        this.font = font;
    }

    public static Integer getFontByTitle(String title){
        for (TitleFontEnum e : TitleFontEnum.values()) {
            if (title.equals(e.getTitle())) {
                return e.getFont();
            }
        }
        return null;
    }
}
