package com.zhuozhengsoft.samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.util.Map;

@RestController
@RequestMapping(value = "/SaveFirstPageAsImg/")
public class SaveFirstPageAsImgController {
    private String dir = ResourceUtils.getURL("classpath:").getPath() + "static\\doc\\";

    public SaveFirstPageAsImgController() throws FileNotFoundException {
    }

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("保存首页为图片", "SaveFirstAsImg()", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法


        //打开Word文档
        poCtrl.webOpen("/doc/SaveFirstPageAsImg/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("SaveFirstPageAsImg/Word");
        return mv;
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        if (fs.getFileExtName().equals(".jpg")) {
            fs.saveToFile(dir + "SaveFirstPageAsImg\\images\\" + fs.getFileName());
        } else {
            fs.saveToFile(dir + "SaveFirstPageAsImg\\" + fs.getFileName());
        }
        fs.setCustomSaveResult("ok");

        fs.close();
    }

}
