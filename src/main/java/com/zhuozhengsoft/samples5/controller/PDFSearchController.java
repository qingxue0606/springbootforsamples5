package com.zhuozhengsoft.samples5.controller;
import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PDFCtrl;
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
@RequestMapping(value="/PDFSearch/")
public class PDFSearchController {
    private String dir= ResourceUtils.getURL("classpath:").getPath()+"static\\doc\\";
    public PDFSearchController() throws FileNotFoundException {
    }
    @RequestMapping(value="PDF", method= RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String,Object> map){
        PDFCtrl poCtrl = new PDFCtrl(request);
        poCtrl.setServerPage(request.getContextPath()+"/poserver.zz"); //此行必须

// Create custom toolbar
        poCtrl.addCustomToolButton("搜索", "SearchText()", 0);
        poCtrl.addCustomToolButton("搜索下一个", "SearchTextNext()", 0);
        poCtrl.addCustomToolButton("搜索上一个", "SearchTextPrev()", 0);
        poCtrl.addCustomToolButton("实际大小", "SetPageReal()", 16);
        poCtrl.addCustomToolButton("适合页面", "SetPageFit()", 17);
        poCtrl.addCustomToolButton("适合宽度", "SetPageWidth()", 18);

        //打开Word文档
        poCtrl.webOpen("/doc/PDFSearch/test.pdf");
        map.put("pageoffice",poCtrl.getHtmlCode("PDFCtrl1"));
        ModelAndView mv = new ModelAndView("PDFSearch/Word");
        return mv;
    }



}
