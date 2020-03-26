package vip.aquan.worddemo.controller;

import com.deepoove.poi.XWPFTemplate;
import org.apache.commons.beanutils.BeanUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import vip.aquan.worddemo.entity.User;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.util.Map;

/**
 * 在这里填描述
 *
 * @author wcp
 * @date 2019/10/31
 */
@Controller
public class WordController {
    /**
     * 导出word，请求exportWord即可
     * @param request
     * @param response
     * @throws Exception
     */
    @RequestMapping(value = "exportWord")
    public void exportWord( HttpServletRequest request, HttpServletResponse response) throws Exception {
        User user = new User();
        user.setName("嘴多多");
        user.setAge("22");
        Map dataMap = BeanUtils.describe(user);

        String fileName = user.getName()+"日常记录报告";

        String agent = request.getHeader("User-Agent");
        boolean isMSIE = ((agent != null && agent.indexOf("MSIE") != -1 ) || ( null != agent && -1 != agent.indexOf("like Gecko")));
        /**旧版本ie直接判断MSIE即可，但是新版本ie跟edge使用了新的内核。msie判断无效，打印agent，先做截取处理。 但是chrome打印出来也包含like Gecko，所以这里chrome等浏览器打印出来也会走gbk的编码，但是不会出现乱码。解决。**/
        if(isMSIE){
            //ie内核
            fileName = new String(fileName.getBytes("GBK"),"ISO8859-1");
        } else {
            //chrome
            fileName = new String(fileName.getBytes("UTF8"), "ISO8859-1");
        }
        response.setHeader("content-disposition", "attachment;filename="+ fileName + ".docx");
//        String realPath = request.getSession().getServletContext().getRealPath("static/daily_record_tamplate.docx");
        String realPath = request.getServletContext().getRealPath("daily_record_tamplate.docx");
        XWPFTemplate template = XWPFTemplate.compile(realPath).render(dataMap);
//        XWPFTemplate template = XWPFTemplate.compile("daily_record_tamplate.docx").render(dataMap);
        //自己指定路径
//        FileOutputStream out = new FileOutputStream("out_template.docx");
        //在网页下载
        OutputStream out = response.getOutputStream();
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
