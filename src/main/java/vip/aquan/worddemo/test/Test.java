package vip.aquan.worddemo.test;

import com.deepoove.poi.XWPFTemplate;

import java.io.FileOutputStream;
import java.util.HashMap;

/**
 * 在这里填描述
 *
 * @author wcp
 * @date 2019/10/31
 */
public class Test {
    public static void main(String[] args) throws Exception {
        XWPFTemplate template = XWPFTemplate.compile("D:\\wcp\\IdeaProject\\demo\\导出word\\WordDemo\\src\\main\\resources\\static\\daily_record_tamplate.docx").render(new HashMap<String, Object>(){{
            put("studentName","王大强a");
        }});
        FileOutputStream out = new FileOutputStream("out_template.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
