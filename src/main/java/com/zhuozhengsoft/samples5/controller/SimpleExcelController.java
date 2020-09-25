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

import static org.springframework.util.ResourceUtils.*;

@RestController
@RequestMapping(value = "/SimpleExcel/")
public class SimpleExcelController {
    private String dir = getURL("classpath:").getPath() + "static/doc/";

    public SimpleExcelController() throws FileNotFoundException {
    }

    @RequestMapping(value = "Excel", method = RequestMethod.GET)
    public ModelAndView showExcel(HttpServletRequest request, Map<String, Object> map) {

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("关闭", "Close", 21);
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开word
        poCtrl.webOpen("/doc/SimpleExcel/test.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));

        ModelAndView mv = new ModelAndView("SimpleExcel/Excel");
        return mv;
    }

    @RequestMapping("save")
    public void saveFile(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "SimpleExcel/" + fs.getFileName());
        fs.close();
    }
}
