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
@RequestMapping(value="/TitleText/")
public class TitleTextController {

    @RequestMapping(value="Word", method= RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String,Object> map){
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath()+"/poserver.zz");//设置服务页面
        //设置PageOffice控件标题栏内容
        poCtrl.setCaption("PageOfficeCtrl控件标题栏内容");
        //打开Word文档
        poCtrl.webOpen("/doc/TitleText/test.doc", OpenModeType.docNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("TitleText/Word");
        return mv;
    }

}